Attribute VB_Name = "basM46V102"
Option Explicit
Private mDemaisObservacoes() As String

Private lNumPagina As Integer
Private lTotPagina As Integer
Private PodeGerarCYX As Boolean
Public gCorretorComIntegracao As Boolean
Private lCountPage As Integer
Private lListaArq As String

Private mEndConfig As String
Private mEndArq As String
Private mEndImagem As String
Private mListaCobertFranqEsp As String
Private mNomForm As String
Private lQtdPaginas As Integer

Public Function gfVerCorrIntegracao() As Boolean
    Dim lCodCorretor2 As Long           'Código do corretor - instalação
    Dim lEncArquivo As Boolean      'Identifica se a verificação foi feita através do arquivo, caso contrário deverá ser buscado do código
    Dim lIndTransmissao As String
    Dim lEncontrou As Boolean     'Identifica se foi encontrato o [grupos] no arquivo
    Dim lEndP366CORR As String        'Endereço do arquivo P366CORR.EXE.
    Dim lNumP366CORR As Integer       'Utilizado com FreeFile para o arquivo P366CORR.EXE.
    Dim lRegistro As String     'Registro.

    'A comparação no arquivo é feita com o número do corretor sem os dois ultimos digitos,
    'os dois ultimos digitos variam mas dentro de um mesmo grupo de corretor.
    'No arquivo P366CORR.EXE não tem o número completo do corretor, está sem os dois últimos
    'dígitos.

    If gM46V999.gTipUsuario <> e_Tipo_Usu_Corretor Then
        gfVerCorrIntegracao = False
        Exit Function
    End If
    lEndP366CORR = gM46V999.gAppPath & "P366CORR.EXE"
    If Not gM46V999.gfPreenchido(Dir(lEndP366CORR)) Then
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
                If gM46V999.gfPreenchido(lRegistro) Then
                    lRegistro = Trim$(lRegistro)
                    If UCase(lRegistro) = "[INTEGRACAO]" Then
                        lEncontrou = True
                    End If

                    If lEncontrou = True Then
                        Select Case Len(lRegistro)
                        Case 8
                            lCodCorretor2 = Left$(lRegistro, 7)
                            lIndTransmissao = "N"
                        Case 9
                            lCodCorretor2 = Left$(lRegistro, 7)
                            lIndTransmissao = Mid$(lRegistro, 9, 1)
                        Case Else
                            lCodCorretor2 = "0000000"
                            lIndTransmissao = "N"
                        End Select
                    End If
                    If lCodCorretor2 = gM46V999.gCod_Corretor Then
                        Exit Do
                    Else
                        lIndTransmissao = "N"
                    End If
                End If
            Loop
            Close lNumP366CORR
        End If
    End If

    If lIndTransmissao = "S" Then
        gfVerCorrIntegracao = True
    Else
        gfVerCorrIntegracao = False
    End If

End Function
Public Sub gpCriarArquivoMicroMassi(pobjstcPed As stcA46V702B, pobjstcPedLoc As stcA46V708B, pFormulario As String, pPagDe As Integer, pPagAte As Integer)
    Dim lNumArqEntr As Integer      'Número do arquivo.
    Dim lNumArqSaida As Integer       'Número do arquivo.
    Dim lNovoArq As String
    Dim lRegistro As String
    Dim lArquivo As String
    Dim lControle As Integer
    Dim lExcluir As Boolean
    Dim lCaminho As String

    If Not gCorretorComIntegracao Then Exit Sub

    If gM46V999.gfPreenchido(gM46V999.gEndSYASUpDt) = True Then
        'Caminho ods arquivos de atualização do SYAS
        'Informado no arquivo SYASPORD.INI
        lCaminho = gM46V999.gEndSYASUpDt
    Else
        'Endereço das bases de dados tarifarias
        lCaminho = gM46V999.gEndServidor
    End If
    If gM46V999.gfPreenchido(lCaminho) = False Then
        lCaminho = gM46V999.gAppPath
    End If

    If Dir(lCaminho & "INTEGRACAO", vbDirectory) = "" Then
        MkDir (lCaminho & "INTEGRACAO")
    End If
    lExcluir = False
    lArquivo = lCaminho & "INTEGRACAO\" & pobjstcPed.NOSSO_NUMERO & "_" & pFormulario & ".txt"
    If Dir(lArquivo) <> "" Then
        Kill lArquivo
    End If
    lNumArqEntr = FreeFile

    Open lArquivo For Output Access Write Lock Read Write As #lNumArqEntr

    Select Case pFormulario
    Case "MULTPRO1"
        Print #lNumArqEntr, pFormulario & " 31/07/2008"
        If Not mfCriDados_MULTPRO1(lNumArqEntr, pobjstcPed, pobjstcPedLoc) Then Exit Sub
    Case "MULTPRO2"
        Print #lNumArqEntr, pFormulario & " 27/08/2009"
        If Not mfCriDados_MULTPRO2(lNumArqEntr, pobjstcPed, pobjstcPedLoc) Then Exit Sub
    Case "MULTPRO3"
        Print #lNumArqEntr, pFormulario & " 31/07/2008"
        If Not mfCriDados_MULTPRO3(lNumArqEntr, pobjstcPed, pobjstcPedLoc, pPagAte) Then Exit Sub
    Case "MULTPRO4"
        Print #lNumArqEntr, pFormulario & " 27/08/2009"
        If Not mfCriDados_MULTPRO4(lNumArqEntr, pobjstcPed, pPagDe, pPagAte) Then Exit Sub
    Case "MULTPROB"
        Print #lNumArqEntr, pFormulario & " 09/08/2007"
        If Not mfCriDados_MULTPROB(lNumArqEntr, pobjstcPed) Then
            lExcluir = True
        End If
    Case "MULTPROC"
        Print #lNumArqEntr, pFormulario & " 09/08/2007"
        If Not mfCriDados_MULTPROC(lNumArqEntr, pobjstcPed) Then
            lExcluir = True
        End If
    Case "MULTPROD"
        Print #lNumArqEntr, pFormulario & " 24/06/2009"
        If Not mfCriDados_MULTPROD(lNumArqEntr, pobjstcPed) Then
            lExcluir = True
        End If
    End Select

    On Error Resume Next
    Close #lNumArqEntr
    If lExcluir = True Then
        Kill lArquivo
    End If

End Sub

Public Function gpGeraFiguraArqWord(pArquivo As String, ByRef pRetorno As Byte, Optional ByVal pOrigem As Integer = 0) As String
    Dim lAplicacao As Object
    Dim lDocumento As Object
    Dim lArquivo As String
    Dim lArray As Variant
    Dim lI As Integer
    Dim lListaArquivos As String

    gpGeraFiguraArqWord = Empty

    On Error GoTo ControleErro
    
    lArquivo = Dir(gM46V999.gAppPath & Replace(pArquivo, ".WMF", "*.WMF"))
    Do While lArquivo <> ""
        lListaArquivos = lListaArquivos & lArquivo & ";"
        lArquivo = Dir()
    Loop
    lArray = Split(lListaArquivos, ";")

    pOrigem = gM46V999.CDblx(pOrigem)

    'Retirado por Simony 05/12/2007
    'If UBound(lArray) > 1 Then
    '    Call mpGeraFiguraArqWordSeparado(pArquivo, pRetorno)
    'End If

    If UBound(lArray) > 0 Then
        pRetorno = 0
        Set lAplicacao = CreateObject("Word.Application")

        lAplicacao.Documents.Add
        If pOrigem <> 1 Then
            lAplicacao.Visible = True
        End If

        Set lDocumento = lAplicacao.Documents(1)

        lDocumento.PageSetup.LineNumbering.Active = False
        lDocumento.PageSetup.TopMargin = lAplicacao.CentimetersToPoints(0.1)
        lDocumento.PageSetup.BottomMargin = lAplicacao.CentimetersToPoints(0.1)
        lDocumento.PageSetup.LeftMargin = lAplicacao.CentimetersToPoints(1)
        lDocumento.PageSetup.RightMargin = lAplicacao.CentimetersToPoints(0.1)

        'call gM46V999.gpGraLog(0,UBound(lArray) & " folhas.")
        For lI = (UBound(lArray) - 1) To 0 Step -1
            If UCase(lArray(lI)) Like "*ITENS*" Then
                lDocumento.PageSetup.Orientation = 1
            End If

            lDocumento.InlineShapes.AddPicture filename:=gM46V999.gAppPath & lArray(lI), _
                                               LinkToFile:=False, SaveWithDocument:=True
        Next lI

        lDocumento.Protect password:="yasuda21", NoReset:=False, Type:=2
        If Dir(gM46V999.gAppPath & Replace(pArquivo, "WMF", "DOC")) <> "" Then
            Kill gM46V999.gAppPath & Replace(pArquivo, "WMF", "DOC")
        End If
  
        If gfIsMetaframe() Then
            gpGeraFiguraArqWord = "L:\" & Replace(pArquivo, "WMF", "DOC")
        Else
            gpGeraFiguraArqWord = gM46V999.gAppPath & Replace(pArquivo, "WMF", "DOC")
        End If

        lDocumento.SaveAs gpGeraFiguraArqWord

        If pOrigem <> 1 Then
            lAplicacao.WindowState = 1
        Else
            lAplicacao.Documents.Close
            lAplicacao.Quit
        End If

        Set lAplicacao = Nothing
        Set lDocumento = Nothing

        For lI = (UBound(lArray) - 1) To 0 Step -1
            Kill gM46V999.gAppPath & lArray(lI)
        Next lI

    End If
    pRetorno = 1
    
    Exit Function
    
ControleErro:
    Call gM46V999.gpGraLog(1, "gpGeraFiguraArqWord - 5156 - " & Err.Description & " - " & Error, vbCritical)
    Set lAplicacao = Nothing
    Set lDocumento = Nothing
End Function
Private Sub mfCriDados_Cabec_ESP_LMI(ByRef pobjstcPed As stcA46V702B, ByRef pLinCXY As Long, ByRef pPosY As Double, pnumFileDados As Integer, pNumCXY As Integer, pPrimeiraFolha As Boolean)
    Dim lAuxiliar1 As Double    'Auxiliar - utilizado apenas como exemplo.
    Dim lAuxiliar2 As String    'Auxiliar - utilizado apenas como exemplo.
    Dim lobjstcPedLoc As stcA46V708B    'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
    Dim lObjstcCorr As stcA46V705B    'Objeto stcA46V705B  (corretor).
    Dim lQuaDias As Integer

    pLinCXY = 1

    If PodeGerarCYX Then
        Print #pNumCXY, "R Orientacao=Retrato; Papel=A4 210 x 297 mm; Altura=29.70; Largura=21.00; Vazio=0.00; ImgY=0.50; ImgX=0.70; ImgAltura=29.70; ImgLargura=21.00"
    End If
    pPosY = 1
    '002. Título do seguro (de acordo com a base de dados).
    '     Neste caso deve ser colocado YASUDA RESIDENCIAL, YASUDA CONDOMÍNIO ou YASUDA EMPRESARIAL.
    lAuxiliar2 = " "
    Select Case pobjstcPed.Cod_Ramo
    Case e_CodRamo_Condominio
        lAuxiliar2 = "YASUDA CONDOMÍNIO"
    Case e_CodRamo_Empresarial
        lAuxiliar2 = "YASUDA EMPRESARIAL"
    Case e_CodRamo_Residencial
        lAuxiliar2 = "YASUDA RESIDENCIAL"
    End Select
    Print #pnumFileDados, lAuxiliar2
    GeraCYX_LMI pLinCXY, "Titulo", pPosY, 14.4, True, "S", pNumCXY

    '004. Processo Susep.
    lAuxiliar2 = " "
    Select Case pobjstcPed.Cod_Ramo
    Case e_CodRamo_Condominio
        lAuxiliar2 = "Processo Susep 15414.004604/2004-72"
    Case e_CodRamo_Empresarial
        lAuxiliar2 = "Processo Susep 15414.004605/2004-17"
    Case e_CodRamo_Residencial
        lAuxiliar2 = "Processo Susep 15414.003659/2004-65"
    End Select

    Print #pnumFileDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    GeraCYX_LMI pLinCXY, "Processo Susep", pPosY, 13, False, "S", pNumCXY

    Set lobjstcPedLoc = pobjstcPed.ColecaoPedLoc.Item(1)

    If pPrimeiraFolha Then
        pPosY = 4.25
    Else
        pPosY = 3.48
    End If
    '006. Número do PI
    lAuxiliar1 = gM46V111.gNum_Pi
    Print #pnumFileDados, Format$(lAuxiliar1, "000000")
    GeraCYX_LMI pLinCXY, "Referencia", pPosY, 4.5, True, "N", pNumCXY

    '007. Nome do segurado
    Print #pnumFileDados, frmT46V101A.objstcPedido.Nom_Segurado
    GeraCYX_LMI pLinCXY, "Nome do segurado", pPosY, 4.5, True, "N", pNumCXY

    '010. Corretor (código e nome) .
    If frmT46V101A.objstcPedido.ColecaoPedCorr.Count > 0 Then
        Set lObjstcCorr = frmT46V101A.objstcPedido.ColecaoPedCorr.Item(1)
        lAuxiliar1 = lObjstcCorr.Cod_Corr
        lAuxiliar2 = lObjstcCorr.gNom_Corretor
    Else
        lAuxiliar1 = 0
        lAuxiliar2 = " "
    End If

    Print #pnumFileDados, lAuxiliar1 & "-" & lAuxiliar2
    GeraCYX_LMI pLinCXY, "Corretor", pPosY, 4.5, True, "N", pNumCXY

    Print #pnumFileDados, gM46V999.gfForData(2, frmT46V101A.objstcPedido.Dat_Ini_Vig)
    GeraCYX_LMI pLinCXY, "Início de vigência", pPosY, 8.5, True, "N", pNumCXY

    Print #pnumFileDados, gM46V999.gfForData(2, frmT46V101A.objstcPedido.Dat_Ter_Vig)
    GeraCYX_LMI pLinCXY, "Término de vigência", pPosY, 8.5, True, "N", pNumCXY

    Exit Sub

ControleErro:
    Call gM46V999.gpGraLog(1, "mfCriDados_ESP_LMI ERRO - " & Err & " - " & Error)

End Sub


Private Sub mfCriDados_Cabec_LMI(ByRef pobjstcPed As stcA46V702B, ByRef pLinCXY As Long, ByRef pPosY As Double, pnumFileDados As Integer, pNumCXY As Integer, Optional ByVal pImpTipoEmissao As Boolean = True)    'Ficha 1000382 - Ajustar cotação de LMI
    Dim lAuxiliar1 As Double    'Auxiliar - utilizado apenas como exemplo.
    Dim lAuxiliar2 As String    'Auxiliar - utilizado apenas como exemplo.
    Dim lobjstcPedLoc As stcA46V708B    'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
    Dim lObjstcCorr As stcA46V705B    'Objeto stcA46V705B  (corretor).
    Dim lQuaDias As Integer

    pLinCXY = 1
    pPosY = 0.5

    If PodeGerarCYX Then
        Print #pNumCXY, "R Orientacao=Retrato; Papel=A4 210 x 297 mm; Altura=29.70; Largura=21.00; Vazio=0.00; ImgY=0.70; ImgX=0.50; ImgAltura=29.70; ImgLargura=21.00"
    End If

    '001. Data Emissão.
    Print #pnumFileDados, Format$(Now, "dd/mm/yyyy")
    GeraCYX_LMI pLinCXY, "Data de Emissao", pPosY, 18.1, True, "N", pNumCXY

    '002. Título do seguro (de acordo com a base de dados).
    '     Neste caso deve ser colocado YASUDA RESIDENCIAL, YASUDA CONDOMÍNIO ou YASUDA EMPRESARIAL.
    lAuxiliar2 = " "
    Select Case pobjstcPed.Cod_Ramo
    Case e_CodRamo_Condominio
        lAuxiliar2 = "YASUDA CONDOMÍNIO"
    Case e_CodRamo_Empresarial
        lAuxiliar2 = "YASUDA EMPRESARIAL"
    Case e_CodRamo_Residencial
        lAuxiliar2 = "YASUDA RESIDENCIAL"
    End Select
    Print #pnumFileDados, lAuxiliar2
    pPosY = 1.1
    GeraCYX_LMI pLinCXY, "Titulo", pPosY, 8.5, False, "S", pNumCXY

    '003. Hora Emissão.
    Print #pnumFileDados, Format$(Now, "hh:nn:ss")
    GeraCYX_LMI pLinCXY, "Hora de Emissao", pPosY, 18.1, True, "N", pNumCXY

    '004. Processo Susep.
    lAuxiliar2 = " "
    Select Case pobjstcPed.Cod_Ramo
    Case e_CodRamo_Condominio
        lAuxiliar2 = "Processo Susep 15414.004604/2004-72"
    Case e_CodRamo_Empresarial
        lAuxiliar2 = "Processo Susep 15414.004605/2004-17"
    Case e_CodRamo_Residencial
        lAuxiliar2 = "Processo Susep 15414.003659/2004-65"
    End Select

    Print #pnumFileDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    GeraCYX_LMI pLinCXY, "Processo Susep", pPosY, 7.5, False, "S", pNumCXY

    Set lobjstcPedLoc = pobjstcPed.ColecaoPedLoc.Item(1)

    '005. Validade.
    If lobjstcPedLoc.Tip_Emissao = e_TipEmissao_RenYasuda Then
        Print #pnumFileDados, Format$(CDate(gM46V999.gfForData(2, pobjstcPed.Dat_Pri_Calculo)) + 30, "dd/mm/yyyy")
    Else
        Print #pnumFileDados, Format$(CDate(gM46V999.gfForData(2, pobjstcPed.Dat_Pri_Calculo)) + 15, "dd/mm/yyyy")
    End If
    GeraCYX_LMI pLinCXY, "Validade", pPosY, 18.1, False, "N", pNumCXY

    pPosY = 3.4
    '006. Número do pedido (Referencia) .
    lAuxiliar1 = pobjstcPed.Num_Ped
    Print #pnumFileDados, Format$(lAuxiliar1, "000000")
    GeraCYX_LMI pLinCXY, "Referencia", pPosY, 4.5, False, "N", pNumCXY

    '007. Tarifa
    lAuxiliar1 = pobjstcPed.Cod_Produto
    If pobjstcPed.Cod_Sequencia = 1 Then
        lAuxiliar2 = lAuxiliar2 & " (A)"
    End If
    Print #pnumFileDados, Format$(lAuxiliar1, "0000")
    GeraCYX_LMI pLinCXY, "Tarifa", pPosY, 18.1, True, "N", pNumCXY

    '008. Nosso Número.
    lAuxiliar2 = pobjstcPed.NOSSO_NUMERO

    If pobjstcPed.objstcCtrlEmis.Tip_Cotac = 2 And pobjstcPed.objstcCtrlEmis.Num_Cotac <> 0 Then
        lAuxiliar2 = lAuxiliar2 & Space(20) & "Siscota: " & pobjstcPed.objstcCtrlEmis.Num_Cotac & Space(30)
    Else
        lAuxiliar2 = lAuxiliar2 & Space(80)
    End If

    'Ficha 1000382 - Ajustar cotação de LMI
    'Comissao e Desconto
    If pobjstcPed.ColecaoPedCorr.Count > 0 Then
        Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(1)
        lAuxiliar1 = lObjstcCorr.Per_Corr
        If pobjstcPed.ColecaoPedCorr.Count > 1 Then
            Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(2)
            lAuxiliar1 = lAuxiliar1 + lObjstcCorr.Per_Corr
            If pobjstcPed.ColecaoPedCorr.Count > 2 Then
                Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(3)
                lAuxiliar1 = lAuxiliar1 + lObjstcCorr.Per_Corr
            End If
        End If
    End If
    lAuxiliar2 = lAuxiliar2 & Format(lAuxiliar1 * 100, "0000")
    lAuxiliar2 = lAuxiliar2 & Format(pobjstcPed.objstcCtrlEmis.Per_Desc_Coml * 100, "0000")


    Print #pnumFileDados, Format$(lAuxiliar2, String(20, "0"))
    GeraCYX_LMI pLinCXY, "Nosso Numero", pPosY, 4.5, True, "N", pNumCXY

    GeraCYX_LMI pLinCXY, "Linha Vazia", pPosY, 0.5, True, "N", pNumCXY

    '009. Segurado
    Print #pnumFileDados, pobjstcPed.Nom_Segurado
    GeraCYX_LMI pLinCXY, "Nome Segurado", pPosY, 4.5, True, "N", pNumCXY

    '010. Corretor (código e nome) .
    If pobjstcPed.ColecaoPedCorr.Count > 0 Then
        Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(1)
        lAuxiliar1 = lObjstcCorr.Cod_Corr
        lAuxiliar2 = lObjstcCorr.gNom_Corretor
    Else
        lAuxiliar1 = 0
        lAuxiliar2 = " "
    End If

    Print #pnumFileDados, Format$(lAuxiliar1, "0000") & " - " & lAuxiliar2
    GeraCYX_LMI pLinCXY, "Corretor", pPosY, 4.5, True, "N", pNumCXY

    '011. Prazo.
    lQuaDias = gM46V999.gfDifDatas(pobjstcPed.Dat_Ini_Vig, pobjstcPed.Dat_Ter_Vig)
    If lQuaDias = 365 Or lQuaDias = 366 Then
        lAuxiliar2 = "Um ano"
    Else
        lAuxiliar2 = lQuaDias & " dias "
    End If

    Print #pnumFileDados, lAuxiliar2
    GeraCYX_LMI pLinCXY, "Prazo", pPosY, 4.5, True, "N", pNumCXY

    'Ficha 1000382 - Ajustar cotação de LMI
    If pImpTipoEmissao = True Then
        'Tipo de emissão .
        lAuxiliar2 = pobjstcPed.ColecaoPedLoc(1).gDesc_TipEmissao
        Print #pnumFileDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        GeraCYX_LMI pLinCXY, "Tipo de Emissao", pPosY, 4.5, False, "N", pNumCXY

        If pobjstcPed.ColecaoPedLoc(1).Tip_Emissao = e_TipEmissao_RenYasuda Or _
           pobjstcPed.ColecaoPedLoc(1).Tip_Emissao = e_TipEmissao_Endosso Then
            'Apólice Anterior
            lAuxiliar2 = pobjstcPed.ColecaoPedLoc(1).Num_Apol_Anterior
            If lAuxiliar2 = "" Then
                lAuxiliar2 = pobjstcPed.Num_Apol_An
            End If
            If lAuxiliar2 <> "" Then
                If pobjstcPed.Cod_Dizer <> 0 Then
                    Print #pnumFileDados, "Apólice:"
                Else
                    Print #pnumFileDados, "Apólice Anterior:"
                End If
                GeraCYX_LMI pLinCXY, "Label Apólice Anterior", pPosY, 8, False, "S", pNumCXY

                Print #pnumFileDados, lAuxiliar2
                GeraCYX_LMI pLinCXY, "Apólice Anterior", pPosY, 11, False, "N", pNumCXY
            End If

            'Item
            lAuxiliar2 = pobjstcPed.ColecaoPedLoc(1).Num_Item_Anterior
            If lAuxiliar2 = "" Then
                lAuxiliar2 = pobjstcPed.ColecaoPedLoc(1).Num_Item
            End If
            If lAuxiliar2 <> "" Then
                Print #pnumFileDados, "Item:"
                GeraCYX_LMI pLinCXY, "Label Item", pPosY, 13, False, "S", pNumCXY

                Print #pnumFileDados, lAuxiliar2
                GeraCYX_LMI pLinCXY, "Item", pPosY, 14, False, "N", pNumCXY
            End If

            'Sinistro
            lAuxiliar2 = IIf(pobjstcPed.ColecaoPedLoc(1).Ind_Sinistro = e_SimNao_Sim, "Sim", "Não")
            If lAuxiliar2 <> "" Then
                Print #pnumFileDados, "Sinistro:"
                GeraCYX_LMI pLinCXY, "Label Sinistro", pPosY, 15, False, "S", pNumCXY

                Print #pnumFileDados, lAuxiliar2
                GeraCYX_LMI pLinCXY, "Sinistro", pPosY, 16.6, False, "N", pNumCXY
            End If
        End If

    End If

    GeraCYX_LMI pLinCXY, "Linha Vazia", pPosY, 0.5, True, "N", pNumCXY
    GeraCYX_LMI pLinCXY, "Linha Vazia", pPosY, 0.5, True, "N", pNumCXY

    Exit Sub

ControleErro:
    Call gM46V999.gpGraLog(1, "mfCriDados_Cabec_LMI ERRO - " & Err & " - " & Error)

End Sub

Private Function mfCriDados_REL_LMI(ByVal pNumArqDados As Integer, ByVal pobjstcPed As stcA46V702B, _
                                    ByRef pEndConfig As String) As Boolean

'Função: cria arquivo de dados - cotação coletiva - capa.

'Parâmetro de entrada...pNumArqDados...Número do arquivo de dados.

    Dim lAuxiliar2 As String    'Auxiliar - utilizado apenas como exemplo.
    Dim i As Integer             'Utilizado com For...Next.
    Dim lMensagem As String
    Dim vCobert() As String
    Dim vArrayItens() As String
    Dim vArrayCoberts() As String
    Dim lPosArray As Integer
    Dim lobjstcPedLoc As stcA46V708B    'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
    Dim lobjstcPedCobert As stcA46V704B    'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)

    Dim lQtdCobert As Integer
    Dim lAchou As Boolean
    Dim lBarra As String * 1

    Dim lPosY As Double
    Dim lNumArqCXY As Integer
    Dim lLinCXY As Long

    Dim QtdItensImpresso As Integer
    Dim MaxItensImpressao As Integer

    On Error GoTo ControleErro

    'ROTINA CRIA CXY

    If gfIsMetaframe() Then
        mEndConfig = "L:\SYAS.CXY"
        mEndArq = "L:\SYAS.DDD"
        Call gM46V999.gpGraLog(0, "Eh Metaframe forcou o L " & mEndConfig)
    Else
        mEndConfig = gM46V999.gAppPath & "SYAS.CXY"
        mEndArq = gM46V999.gAppPath & "SYAS.DDD"
    End If

    lListaArq = mEndArq & "|" & mEndConfig & "|" & gM46V999.gEndImagRelLMI & ";"


    PodeGerarCYX = True
    lNumArqCXY = FreeFile

    On Error Resume Next
    Open mEndConfig For Output Access Write Lock Read Write As #lNumArqCXY

    mfCriDados_Cabec_LMI pobjstcPed, lLinCXY, lPosY, pNumArqDados, lNumArqCXY, False    'Ficha 1000382 - Ajustar cotação de LMI

    lQtdCobert = 0
    'Carrega um array com as cobertura
    For Each lobjstcPedLoc In pobjstcPed.ColecaoPedLoc
        For Each lobjstcPedCobert In lobjstcPedLoc.ColecaoPedCobert
            lAchou = False
            For i = 1 To lQtdCobert
                If vCobert(0, i) = lobjstcPedCobert.Cod_Cobert Then
                    lAchou = True
                    Exit For
                End If
            Next
            If Not lAchou Then
                lQtdCobert = lQtdCobert + 1
                ReDim Preserve vCobert(1, lQtdCobert)
                vCobert(0, lQtdCobert) = lobjstcPedCobert.Cod_Cobert
                vCobert(1, lQtdCobert) = lobjstcPedCobert.gNomeCobertura
            End If
        Next
    Next

    lPosArray = 0

    If lQtdCobert >= 13 Then
        MaxItensImpressao = 2
    ElseIf lQtdCobert >= 6 Then
        MaxItensImpressao = 4
    Else
        MaxItensImpressao = 6
    End If

    ReDim Preserve vArrayItens(1, 7)
    ReDim Preserve vArrayCoberts(1, lQtdCobert, 1)

    For Each lobjstcPedLoc In pobjstcPed.ColecaoPedLoc
        lBarra = "|"
        'Ficha 1000576
        vArrayItens(lPosArray, 0) = "ITEM: " & lobjstcPedLoc.Num_Item & gfPreencheBranco(Format(IIf(lobjstcPedLoc.Insp_Obr, "INSP", "")), 75)
        lAuxiliar2 = ""
        lAuxiliar2 = lAuxiliar2 & Trim(lobjstcPedLoc.Cod_Tip_Logr_Risco) & " "
        lAuxiliar2 = lAuxiliar2 & Trim(Left(lobjstcPedLoc.Nom_Logr_Risco, 30)) & ",  "
        lAuxiliar2 = lAuxiliar2 & Trim(lobjstcPedLoc.objstcPedInsp.Num_Logr_Risco)
        lAuxiliar2 = lAuxiliar2 & IIf(Trim(lobjstcPedLoc.Dsc_Compl_Risco) = "", "", " - " & Trim(lobjstcPedLoc.Dsc_Compl_Risco)) & " "
        vArrayItens(lPosArray, 1) = "LOCAL: " & lAuxiliar2
        lAuxiliar2 = ""
        lAuxiliar2 = lAuxiliar2 & Mid(Trim(lobjstcPedLoc.Nom_Bairro_Risco), 1, 15) & " - "
        lAuxiliar2 = lAuxiliar2 & Trim(lobjstcPedLoc.Nom_Cid_Risco) & " "
        lAuxiliar2 = lAuxiliar2 & Trim(lobjstcPedLoc.Sig_UF_Risco) & " "
        lAuxiliar2 = lAuxiliar2 & Format(lobjstcPedLoc.Num_Cep_Risco, "00000-000")
        vArrayItens(lPosArray, 2) = lAuxiliar2
        vArrayItens(lPosArray, 3) = "Rubrica: " & lobjstcPedLoc.Cod_Clasf & "." & lobjstcPedLoc.Cod_Compl_Clasf
        vArrayItens(lPosArray, 4) = "Valor do Risco: R$"
        vArrayItens(lPosArray, 5) = gfPreencheBranco(Format(lobjstcPedLoc.Val_Vr, "standard"), 20)
        vArrayItens(lPosArray, 6) = "LMI do Local R$"

        For i = 1 To lQtdCobert
            vArrayCoberts(lPosArray, i, 0) = vCobert(1, i)
            vArrayCoberts(lPosArray, i, 1) = 0
            For Each lobjstcPedCobert In lobjstcPedLoc.ColecaoPedCobert
                If vCobert(0, i) = lobjstcPedCobert.Cod_Cobert Then
                    vArrayCoberts(lPosArray, i, 0) = IIf(Trim(UCase(lobjstcPedCobert.Compl_Desc_Cobertura)) = "PREDIO", "(P)", IIf(Trim(UCase(lobjstcPedCobert.Compl_Desc_Cobertura)) = "CONTEUDO", "(C)", Empty)) & vArrayCoberts(lPosArray, i, 0)
                    vArrayCoberts(lPosArray, i, 1) = lobjstcPedCobert.Val_IS
                    Exit For
                End If
            Next
        Next
        If lPosArray = 0 Then
            lPosArray = 1
        Else
            If QtdItensImpresso >= MaxItensImpressao Then
                GoSub Rodape
                mfCriDados_Cabec_LMI pobjstcPed, lLinCXY, lPosY, pNumArqDados, lNumArqCXY, False    'Ficha 1000382 - Ajustar cotação de LMI
                QtdItensImpresso = 0
            End If
            GoSub DescarregarArrays
        End If
    Next
    If Trim(vArrayItens(0, 0)) <> "" Then
        If QtdItensImpresso >= MaxItensImpressao Then
            GoSub Rodape
            mfCriDados_Cabec_LMI pobjstcPed, lLinCXY, lPosY, pNumArqDados, lNumArqCXY, False    'Ficha 1000382 - Ajustar cotação de LMI
            QtdItensImpresso = 0
        End If
        GoSub DescarregarArrays
    End If
    Do While QtdItensImpresso < MaxItensImpressao And QtdItensImpresso <> 0
        GoSub DescarregarArrays    'descarregar array vazia
    Loop

    GoSub Rodape

    'Print #pNumArqDados, pobjstcPed.Cod_User_Liber & " - " & gM46V999.gfForData(2, pobjstcPed.Dat_Liber) & " - " & Mid(Format(pobjstcPed.Hor_Liber, "0000"), 1, 2) & ":" & Mid(Format(pobjstcPed.Hor_Liber, "0000"), 3, 2)

    mfCriDados_REL_LMI = True
    Exit Function

Rodape:
    If PodeGerarCYX Then
        Print #lNumArqCXY, "D Descricao = Rodape Paginas; PosY=28.5; PosX=18; Fonte=ARIAL NARROW;       Tamanho=12; Negrito=N; Italico=N; Sublinhado=N; Riscado=N; Cor=Preto"
        Print #lNumArqCXY, " "
        Close #lNumArqCXY
    End If
    PodeGerarCYX = False
    Print #pNumArqDados, "Página: " & lNumPagina & "/" & lTotPagina
    lNumPagina = lNumPagina + 1

    Return

DescarregarArrays:

    'Numero do Item
    Print #pNumArqDados, vArrayItens(0, 0)
    GeraCYX_LMI lLinCXY, "Descricao do Item", lPosY, 0.5, False, "N", lNumArqCXY
    Print #pNumArqDados, lBarra
    GeraCYX_LMI lLinCXY, "Descricao do Item", lPosY, 10.6, False, "N", lNumArqCXY
    Print #pNumArqDados, IIf(vArrayItens(1, 0) = "", " ", vArrayItens(1, 0))
    GeraCYX_LMI lLinCXY, "Descricao do Item", lPosY, 11, True, "N", lNumArqCXY
    'Endereço
    Print #pNumArqDados, vArrayItens(0, 1)
    GeraCYX_LMI lLinCXY, "Descricao do Item", lPosY, 0.5, False, "N", lNumArqCXY
    Print #pNumArqDados, lBarra
    GeraCYX_LMI lLinCXY, "Descricao do Item", lPosY, 10.6, False, "N", lNumArqCXY
    Print #pNumArqDados, IIf(vArrayItens(1, 1) = "", " ", vArrayItens(1, 1))
    GeraCYX_LMI lLinCXY, "Descricao do Item", lPosY, 11, True, "N", lNumArqCXY
    'Endereço 2
    Print #pNumArqDados, vArrayItens(0, 2)
    GeraCYX_LMI lLinCXY, "Descricao do Item", lPosY, 0.5, False, "N", lNumArqCXY
    Print #pNumArqDados, lBarra
    GeraCYX_LMI lLinCXY, "Descricao do Item", lPosY, 10.6, False, "N", lNumArqCXY
    Print #pNumArqDados, IIf(vArrayItens(1, 2) = "", " ", vArrayItens(1, 2))
    GeraCYX_LMI lLinCXY, "Descricao do Item", lPosY, 11, True, "N", lNumArqCXY
    'Rubrica
    Print #pNumArqDados, vArrayItens(0, 3)
    GeraCYX_LMI lLinCXY, "Descricao do Item", lPosY, 0.5, False, "N", lNumArqCXY
    Print #pNumArqDados, lBarra
    GeraCYX_LMI lLinCXY, "Descricao do Item", lPosY, 10.6, False, "N", lNumArqCXY
    Print #pNumArqDados, IIf(vArrayItens(1, 3) = "", " ", vArrayItens(1, 3))
    GeraCYX_LMI lLinCXY, "Descricao do Item", lPosY, 11, True, "N", lNumArqCXY
    'VR
    Print #pNumArqDados, vArrayItens(0, 4)
    GeraCYX_LMI lLinCXY, "Descricao do Item", lPosY, 0.5, False, "N", lNumArqCXY

    Print #pNumArqDados, gfPreencheBranco(Format(vArrayItens(0, 5), "standard"), 15)
    GeraCYX_LMI lLinCXY, "Descricao do Item", lPosY, 6.5, False, "N", lNumArqCXY, , , "COURIER NEW"
    Print #pNumArqDados, lBarra
    GeraCYX_LMI lLinCXY, "Descricao do Item", lPosY, 10.6, False, "N", lNumArqCXY
    Print #pNumArqDados, IIf(vArrayItens(1, 4) = "", " ", vArrayItens(1, 4))
    GeraCYX_LMI lLinCXY, "Descricao do Item", lPosY, 11, False, "N", lNumArqCXY
    Print #pNumArqDados, gfPreencheBranco(Format(vArrayItens(1, 5), "standard"), 15)
    GeraCYX_LMI lLinCXY, "Descricao do Item", lPosY, 16.7, True, "N", lNumArqCXY, , , "COURIER NEW"

    Print #pNumArqDados, lBarra
    GeraCYX_LMI lLinCXY, "Descricao do Item", lPosY, 10.6, True, "N", lNumArqCXY
    'LMI
    Print #pNumArqDados, vArrayItens(0, 6)
    GeraCYX_LMI lLinCXY, "Descricao do Item", lPosY, 8, False, "N", lNumArqCXY
    Print #pNumArqDados, lBarra
    GeraCYX_LMI lLinCXY, "Descricao do Item", lPosY, 10.6, False, "N", lNumArqCXY
    Print #pNumArqDados, IIf(vArrayItens(1, 6) = "", " ", vArrayItens(1, 6))
    GeraCYX_LMI lLinCXY, "Descricao do Item", lPosY, 18, True, "N", lNumArqCXY
    'Coberturas
    For i = 1 To lQtdCobert

        Print #pNumArqDados, Mid(vArrayCoberts(0, i, 0), 1, 27)
        GeraCYX_LMI lLinCXY, "Descricao do Cobertura", lPosY, 0.5, False, "N", lNumArqCXY
        Print #pNumArqDados, gfPreencheBranco(Format(vArrayCoberts(0, i, 1), "standard"), 15)
        GeraCYX_LMI lLinCXY, "Descricao do Cobertura", lPosY, 6.5, False, "N", lNumArqCXY, , , "COURIER NEW"

        Print #pNumArqDados, lBarra
        GeraCYX_LMI lLinCXY, "Descricao do Cobertura", lPosY, 10.6, False, "N", lNumArqCXY

        Print #pNumArqDados, IIf(vArrayCoberts(1, i, 0) = "", " ", Mid(vArrayCoberts(1, i, 0), 1, 27))
        GeraCYX_LMI lLinCXY, "Descricao do Cobertura", lPosY, 11, False, "N", lNumArqCXY
        Print #pNumArqDados, gfPreencheBranco(Format(vArrayCoberts(1, i, 1), "standard"), 15)
        GeraCYX_LMI lLinCXY, "Descricao do Cobertura", lPosY, 16.7, True, "N", lNumArqCXY, , , "COURIER NEW"
    Next

    GeraCYX_LMI lLinCXY, "Linha Vazia", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed, , True

    'Sempre colocar espaços
    lBarra = " "
    For i = 0 To 6
        For lPosArray = 0 To 1
            vArrayItens(lPosArray, i) = "  "
        Next
    Next
    For i = 1 To lQtdCobert
        For lPosArray = 0 To 1
            vArrayCoberts(lPosArray, i, 0) = "  "
            vArrayCoberts(lPosArray, i, 1) = "  "
        Next
    Next
    QtdItensImpresso = QtdItensImpresso + 2
    lPosArray = 0

    Return

ControleErro:
    Call gM46V999.gpGraLog(1, "mfCriDados_REL_LMI ERRO - " & Err & " - " & Error)

End Function
Private Function mfCriDados_MULTLMIR(ByVal pNumArqDados As Integer, ByVal pobjstcPed As stcA46V702B, _
                                     ByRef pEndConfig As String) As Boolean

'Função: cria arquivo de dados - cotação coletiva - capa.

'Parâmetro de entrada...pNumArqDados...Número do arquivo de dados.

    Dim lAuxiliar2 As String    'Auxiliar - utilizado apenas como exemplo.
    Dim i As Integer             'Utilizado com For...Next.
    Dim lMensagem As String
    Dim vCobert() As String
    Dim vArrayItens() As String
    Dim vArrayCoberts() As String
    Dim lPosArray As Integer
    Dim lobjstcPedLoc As stcA46V708B    'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
    Dim lobjstcPedCobert As stcA46V704B    'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)

    Dim lQtdCobert As Integer
    Dim lAchou As Boolean
    Dim lBarra As String * 1

    Dim lPosY As Double
    Dim lNumArqCXY As Integer
    Dim lLinCXY As Long

    Dim QtdItensImpresso As Integer
    Dim MaxItensImpressao As Integer

    On Error GoTo ControleErro

    'ROTINA CRIA CXY

    If gfIsMetaframe() Then
        mEndConfig = "L:\SYAS.CXY"
        mEndArq = "L:\SYAS.DDD"
        Call gM46V999.gpGraLog(0, "Eh Metaframe forcou o L " & mEndConfig)
    Else
        mEndConfig = gM46V999.gAppPath & "SYAS.CXY"
        mEndArq = gM46V999.gAppPath & "SYAS.DDD"



    End If

    lListaArq = mEndArq & "|" & mEndConfig & "|" & gM46V999.gEndImagESPRelLMI & ";"

    PodeGerarCYX = True
    lNumArqCXY = FreeFile

    On Error Resume Next
    Open mEndConfig For Output Access Write Lock Read Write As #lNumArqCXY
    lPosY = 0.5
    mfCriDados_Cabec_ESP_LMI pobjstcPed, lLinCXY, lPosY, pNumArqDados, lNumArqCXY, False

    lQtdCobert = 0
    'Carrega um array com as cobertura
    For Each lobjstcPedLoc In pobjstcPed.ColecaoPedLoc
        For Each lobjstcPedCobert In lobjstcPedLoc.ColecaoPedCobert
            lAchou = False
            For i = 1 To lQtdCobert
                If vCobert(0, i) = lobjstcPedCobert.Cod_Cobert Then
                    lAchou = True
                    Exit For
                End If
            Next
            If Not lAchou Then
                lQtdCobert = lQtdCobert + 1
                ReDim Preserve vCobert(1, lQtdCobert)
                vCobert(0, lQtdCobert) = lobjstcPedCobert.Cod_Cobert
                vCobert(1, lQtdCobert) = lobjstcPedCobert.gNomeCobertura
            End If
        Next
    Next

    lPosArray = 0

    If lQtdCobert >= 13 Then
        MaxItensImpressao = 2
    ElseIf lQtdCobert >= 6 Then
        MaxItensImpressao = 4
    Else
        MaxItensImpressao = 6
    End If

    lPosY = 6.5
    ReDim Preserve vArrayItens(1, 7)
    ReDim Preserve vArrayCoberts(1, lQtdCobert, 1)

    For Each lobjstcPedLoc In pobjstcPed.ColecaoPedLoc
        lBarra = "|"
        vArrayItens(lPosArray, 0) = "ITEM: " & lobjstcPedLoc.Num_Item
        lAuxiliar2 = ""
        lAuxiliar2 = lAuxiliar2 & Trim(lobjstcPedLoc.Cod_Tip_Logr_Risco) & " "
        lAuxiliar2 = lAuxiliar2 & Trim(Left(lobjstcPedLoc.Nom_Logr_Risco, 30)) & ",  "
        lAuxiliar2 = lAuxiliar2 & Trim(lobjstcPedLoc.objstcPedInsp.Num_Logr_Risco)
        lAuxiliar2 = lAuxiliar2 & IIf(Trim(lobjstcPedLoc.Dsc_Compl_Risco) = "", "", " - " & Trim(lobjstcPedLoc.Dsc_Compl_Risco)) & " "
        vArrayItens(lPosArray, 1) = "LOCAL: " & lAuxiliar2
        lAuxiliar2 = ""
        lAuxiliar2 = lAuxiliar2 & Mid(Trim(lobjstcPedLoc.Nom_Bairro_Risco), 1, 15) & " - "
        lAuxiliar2 = lAuxiliar2 & Trim(lobjstcPedLoc.Nom_Cid_Risco) & " "
        lAuxiliar2 = lAuxiliar2 & Trim(lobjstcPedLoc.Sig_UF_Risco) & " "
        lAuxiliar2 = lAuxiliar2 & Format(lobjstcPedLoc.Num_Cep_Risco, "00000-000")
        vArrayItens(lPosArray, 2) = lAuxiliar2
        vArrayItens(lPosArray, 3) = "Rubrica: " & lobjstcPedLoc.Cod_Clasf & "." & lobjstcPedLoc.Cod_Compl_Clasf
        vArrayItens(lPosArray, 4) = "Valor do Risco: R$ "
        vArrayItens(lPosArray, 5) = gfPreencheBranco(Format(lobjstcPedLoc.Val_Vr, "standard"), 20)
        vArrayItens(lPosArray, 6) = "LMI do Local R$"

        For i = 1 To lQtdCobert
            vArrayCoberts(lPosArray, i, 0) = vCobert(1, i)
            vArrayCoberts(lPosArray, i, 1) = 0
            For Each lobjstcPedCobert In lobjstcPedLoc.ColecaoPedCobert
                If vCobert(0, i) = lobjstcPedCobert.Cod_Cobert Then
                    vArrayCoberts(lPosArray, i, 1) = lobjstcPedCobert.Val_IS
                    Exit For
                End If
            Next
        Next
        If lPosArray = 0 Then
            lPosArray = 1
        Else
            If QtdItensImpresso >= MaxItensImpressao Then
                GoSub Rodape
                mfCriDados_Cabec_ESP_LMI pobjstcPed, lLinCXY, lPosY, pNumArqDados, lNumArqCXY, False
                QtdItensImpresso = 0
            End If
            GoSub DescarregarArrays
        End If
    Next
    If Trim(vArrayItens(0, 0)) <> "" Then
        If QtdItensImpresso >= MaxItensImpressao Then
            GoSub Rodape
            mfCriDados_Cabec_ESP_LMI pobjstcPed, lLinCXY, lPosY, pNumArqDados, lNumArqCXY, False
            QtdItensImpresso = 0
        End If
        GoSub DescarregarArrays
    End If
    Do While QtdItensImpresso < MaxItensImpressao And QtdItensImpresso <> 0
        GoSub DescarregarArrays    'descarregar array vazia
    Loop

    GoSub Rodape

    'Print #pNumArqDados, pobjstcPed.Cod_User_Liber & " - " & gM46V999.gfForData(2, pobjstcPed.Dat_Liber) & " - " & Mid(Format(pobjstcPed.Hor_Liber, "0000"), 1, 2) & ":" & Mid(Format(pobjstcPed.Hor_Liber, "0000"), 3, 2)

    mfCriDados_MULTLMIR = True
    Exit Function

Rodape:
    If PodeGerarCYX Then
        Print #lNumArqCXY, "D Descricao = Rodape Paginas; PosY=28.5; PosX=18; Fonte=ARIAL NARROW;       Tamanho=12; Negrito=N; Italico=N; Sublinhado=N; Riscado=N; Cor=Preto"
        Print #lNumArqCXY, " "
        Close #lNumArqCXY
    End If
    PodeGerarCYX = False
    Print #pNumArqDados, "Página: " & lNumPagina & "/@@TOTALPAG@@"
    lNumPagina = lNumPagina + 1

    Return

DescarregarArrays:

    'Numero do Item
    Print #pNumArqDados, vArrayItens(0, 0)
    GeraCYX_LMI lLinCXY, "Descricao do Item", lPosY, 0.5, False, "N", lNumArqCXY
    Print #pNumArqDados, lBarra
    GeraCYX_LMI lLinCXY, "Descricao do Item", lPosY, 10.6, False, "N", lNumArqCXY
    Print #pNumArqDados, vArrayItens(1, 0)
    GeraCYX_LMI lLinCXY, "Descricao do Item", lPosY, 11, True, "N", lNumArqCXY
    'Endereço
    Print #pNumArqDados, vArrayItens(0, 1)
    GeraCYX_LMI lLinCXY, "Descricao do Item", lPosY, 0.5, False, "N", lNumArqCXY
    Print #pNumArqDados, lBarra
    GeraCYX_LMI lLinCXY, "Descricao do Item", lPosY, 10.6, False, "N", lNumArqCXY
    Print #pNumArqDados, vArrayItens(1, 1)
    GeraCYX_LMI lLinCXY, "Descricao do Item", lPosY, 11, True, "N", lNumArqCXY
    'Endereço 2
    Print #pNumArqDados, vArrayItens(0, 2)
    GeraCYX_LMI lLinCXY, "Descricao do Item", lPosY, 0.5, False, "N", lNumArqCXY
    Print #pNumArqDados, lBarra
    GeraCYX_LMI lLinCXY, "Descricao do Item", lPosY, 10.6, False, "N", lNumArqCXY
    Print #pNumArqDados, vArrayItens(1, 2)
    GeraCYX_LMI lLinCXY, "Descricao do Item", lPosY, 11, True, "N", lNumArqCXY
    'Rubrica
    Print #pNumArqDados, vArrayItens(0, 3)
    GeraCYX_LMI lLinCXY, "Descricao do Item", lPosY, 0.5, False, "N", lNumArqCXY
    Print #pNumArqDados, lBarra
    GeraCYX_LMI lLinCXY, "Descricao do Item", lPosY, 10.6, False, "N", lNumArqCXY
    Print #pNumArqDados, vArrayItens(1, 3)
    GeraCYX_LMI lLinCXY, "Descricao do Item", lPosY, 11, True, "N", lNumArqCXY
    'VR
    Print #pNumArqDados, vArrayItens(0, 4)
    GeraCYX_LMI lLinCXY, "Descricao do Item", lPosY, 0.5, False, "N", lNumArqCXY

    Print #pNumArqDados, gfPreencheBranco(Format(vArrayItens(0, 5), "standard"), 15)
    GeraCYX_LMI lLinCXY, "Descricao do Item", lPosY, 6.5, False, "N", lNumArqCXY, , , "COURIER NEW"
    Print #pNumArqDados, lBarra
    GeraCYX_LMI lLinCXY, "Descricao do Item", lPosY, 10.6, False, "N", lNumArqCXY
    Print #pNumArqDados, vArrayItens(1, 4)
    GeraCYX_LMI lLinCXY, "Descricao do Item", lPosY, 11, False, "N", lNumArqCXY
    Print #pNumArqDados, gfPreencheBranco(Format(vArrayItens(1, 5), "standard"), 15)
    GeraCYX_LMI lLinCXY, "Descricao do Item", lPosY, 16.7, True, "N", lNumArqCXY, , , "COURIER NEW"

    Print #pNumArqDados, lBarra
    GeraCYX_LMI lLinCXY, "Descricao do Item", lPosY, 10.6, True, "N", lNumArqCXY
    'LMI
    Print #pNumArqDados, vArrayItens(0, 6)
    GeraCYX_LMI lLinCXY, "Descricao do Item", lPosY, 8, False, "N", lNumArqCXY
    Print #pNumArqDados, lBarra
    GeraCYX_LMI lLinCXY, "Descricao do Item", lPosY, 10.6, False, "N", lNumArqCXY
    Print #pNumArqDados, vArrayItens(1, 6)
    GeraCYX_LMI lLinCXY, "Descricao do Item", lPosY, 18, True, "N", lNumArqCXY
    'Coberturas
    For i = 1 To lQtdCobert

        Print #pNumArqDados, Mid(vArrayCoberts(0, i, 0), 1, 27)
        GeraCYX_LMI lLinCXY, "Descricao do Cobertura", lPosY, 0.5, False, "N", lNumArqCXY
        Print #pNumArqDados, gfPreencheBranco(Format(vArrayCoberts(0, i, 1), "standard"), 15)
        GeraCYX_LMI lLinCXY, "Descricao do Cobertura", lPosY, 6.5, False, "N", lNumArqCXY, , , "COURIER NEW"

        Print #pNumArqDados, lBarra
        GeraCYX_LMI lLinCXY, "Descricao do Cobertura", lPosY, 10.6, False, "N", lNumArqCXY

        Print #pNumArqDados, Mid(vArrayCoberts(1, i, 0), 1, 27)
        GeraCYX_LMI lLinCXY, "Descricao do Cobertura", lPosY, 11, False, "N", lNumArqCXY
        Print #pNumArqDados, gfPreencheBranco(Format(vArrayCoberts(1, i, 1), "standard"), 15)
        GeraCYX_LMI lLinCXY, "Descricao do Cobertura", lPosY, 16.7, True, "N", lNumArqCXY, , , "COURIER NEW"
    Next

    GeraCYX_LMI lLinCXY, "Linha Vazia", lPosY, 0.5, True, "N", lNumArqCXY

    'Sempre colocar espaços
    lBarra = " "
    For i = 0 To 6
        For lPosArray = 0 To 1
            vArrayItens(lPosArray, i) = "  "
        Next
    Next
    For i = 1 To lQtdCobert
        For lPosArray = 0 To 1
            vArrayCoberts(lPosArray, i, 0) = "  "
            vArrayCoberts(lPosArray, i, 1) = "  "
        Next
    Next
    QtdItensImpresso = QtdItensImpresso + 2
    lPosArray = 0

    Return

ControleErro:
    Call gM46V999.gpGraLog(1, "mfCriDados_REL_LMI ERRO - " & Err & " - " & Error)

End Function

Sub GeraCYX_LMI(ByRef pLinha As Long, pDescricao As String, ByRef pPosY As Double, pPosX As Double, pSomaY As Boolean, pNegrito As String, pnumFile As Integer, Optional pnumFileDados As Integer, Optional ByRef pobjstcPed As stcA46V702B, Optional pFonte As String, Optional pUltimaLinha As Boolean = False)
    Dim lAux As String * 43
    Dim lFonte As String * 15
    On Error GoTo ControleErro
    lFonte = pFonte & ";"
    If IsMissing(pFonte) Or pFonte = "" Then
        lFonte = "ARIAL NARROW;"
    End If
    If pDescricao <> "Linha Vazia" And PodeGerarCYX Then
        lAux = Format(pLinha, "000") & " " & pDescricao
        Print #pnumFile, "D Descricao = " & lAux & "; PosY=" & Replace(Format(pPosY, "00.00"), ",", ".") & "; PosX=" & Replace(Format(pPosX, "00.00"), ",", ".") & "; Fonte=" & lFonte & " Tamanho=12; Negrito=" & pNegrito & "; Italico=N; Sublinhado=N; Riscado=N; Cor=Preto"
        pLinha = pLinha + 1
    End If
    If pSomaY Then
        pPosY = pPosY + 0.5
    End If
    If pPosY > 26.9 And pPosY < 29 And pUltimaLinha = False Then
        If PodeGerarCYX Then
            lAux = "Rodape Paginas"
            Print #pnumFile, "D Descricao = " & lAux & "; PosY=28.5; PosX=18; Fonte=ARIAL NARROW;       Tamanho=12; Negrito=N; Italico=N; Sublinhado=N; Riscado=N; Cor=Preto"
            Print #pnumFile, " "
            Close #pnumFile
            On Error Resume Next
            lCountPage = lCountPage + 1    'Soma 1 pois houve essa quebra no principal

            If gfIsMetaframe() Then
                mEndArq = "L:\" & "SYAS" & lCountPage & ".DDD"
                mEndConfig = "L:\" & "SYAS" & lCountPage & ".CXY"
            Else
                mEndArq = gM46V999.gAppPath & "SYAS" & lCountPage & ".DDD"
                mEndConfig = gM46V999.gAppPath & "SYAS" & lCountPage & ".CXY"
            End If

            If mNomForm = "MULTLMIR" Then
                'Or mNomForm = "MULTLMIE" Then
                lListaArq = lListaArq & mEndArq & "|" & mEndConfig & "|" & gM46V999.gEndImagESPRelLMI & ";"
            ElseIf mNomForm = "MULTLMIE" Then
                'Or mNomForm = "MULTLMIE" Then
                lListaArq = lListaArq & mEndArq & "|" & mEndConfig & "|" & gM46V999.gEndImagESPLMI & ";"
            ElseIf mNomForm = "COT_LMI" Then
                lListaArq = lListaArq & mEndArq & "|" & mEndConfig & "|" & gM46V999.gEndImagCotLMI & ";"
            Else
                lListaArq = lListaArq & mEndArq & "|" & mEndConfig & "|" & gM46V999.gEndImagRelLMI & ";"
            End If
            On Error GoTo ControleErro
            Open mEndConfig For Output Access Write Lock Read Write As #pnumFile

            lTotPagina = lTotPagina + 1    'Soma 1 pois houve essa quebra no principal
            Print #pnumFileDados, "Página: " & lNumPagina & "/@@TOTALPAG@@"
            Close #pnumFileDados

            lNumPagina = lNumPagina + 1
            Open mEndArq For Output Access Write Lock Read Write As #pnumFileDados
            If mNomForm = "MULTLMIR" Or mNomForm = "MULTLMIE" Then
                mfCriDados_Cabec_ESP_LMI pobjstcPed, pLinha, pPosY, pnumFileDados, pnumFile, IIf(mNomForm = "MULTLMIE", True, False)
                pPosY = 7
            Else
                mfCriDados_Cabec_LMI pobjstcPed, pLinha, pPosY, pnumFileDados, pnumFile
            End If
        End If
    End If
    Exit Sub

ControleErro:
    Call gM46V999.gpGraLog(1, "GeraCYX_LMI ERRO - " & Err & " - " & Error)
    Exit Sub

    Return
End Sub
Sub GeraCYX_PROP2(ByRef pLinha As Long, pDescricao As String, ByRef pPosY As Double, pPosX As Double, pSomaY As Boolean, pNegrito As String, pnumFile As Integer, Optional pnumFileDados As Integer, Optional ByRef pobjstcPed As stcA46V702B, Optional ByRef pobjstcPedLoc As stcA46V708B, Optional pFonte As String)
    Dim lAux As String * 43
    Dim lFonte As String * 15



    'On Error GoTo ControleErro

    lFonte = pFonte & ";"
    If IsMissing(pFonte) Or pFonte = "" Then
        lFonte = "ARIAL NARROW;"
    End If
    If pLinha = 43 Then
        On Error Resume Next
        Call mpNovaPag(pLinha, pPosY, pPosX, pnumFile, pnumFileDados, pobjstcPed, pobjstcPedLoc)
    End If
    If pDescricao <> "Linha Vazia" And PodeGerarCYX Then
        lAux = Format(pLinha, "000") & " " & pDescricao
        Print #pnumFile, "D Descricao = " & lAux & "; PosY=" & Replace(Format(pPosY, "00.00"), ",", ".") & "; PosX=" & Replace(Format(pPosX, "00.00"), ",", ".") & "; Fonte=" & lFonte & " Tamanho=12; Negrito=" & pNegrito & "; Italico=N; Sublinhado=N; Riscado=N; Cor=Preto"
        If pSomaY Then
            pLinha = pLinha + 1
        End If
    End If
    If pSomaY Then
        pPosY = pPosY + 0.5
    End If
    Exit Sub

ControleErro:
    Call gM46V999.gpGraLog(1, "GeraCYX_Prop2 ERRO - " & Err & " - " & Error)
    Exit Sub

    Return
End Sub


Private Sub mpNovaPag(pLinha As Long, pPosY As Double, pPosX As Double, pnumFile As Integer, pnumFileDados As Integer, pobjstcPed As stcA46V702B, pobjstcPedLoc As stcA46V708B)
    lCountPage = lCountPage + 1    'Soma 1 pois houve essa quebra no principal
    lQtdPaginas = lQtdPaginas + 1

    If gfIsMetaframe() Then
        mEndArq = "L:\" & "SYAS" & lCountPage & ".DDD"
        mEndConfig = "L:\" & "SYAS" & lCountPage & ".CXY"
    Else
        mEndArq = gM46V999.gAppPath & "SYAS" & lCountPage & ".DDD"
        mEndConfig = gM46V999.gAppPath & "SYAS" & lCountPage & ".CXY"
    End If

    If pobjstcPed.Cod_Tip_Doc = e_TipDoc_Proposta Then
        lListaArq = lListaArq & mEndArq & "|" & mEndConfig & "|" & gM46V999.gEndImagPro2 & ";"
    Else
        lListaArq = lListaArq & mEndArq & "|" & mEndConfig & "|" & gM46V999.gEndImagCotA & ";"
    End If
    Close #pnumFile
    Close #pnumFileDados
    Open mEndConfig For Output Access Write Lock Read Write As #pnumFile

    Open mEndArq For Output Access Write Lock Read Write As #pnumFileDados
    pLinha = 0
    pPosX = 0.6
    pPosY = 5
    If pobjstcPed.Cod_Tip_Doc = e_TipDoc_Proposta Then
        mfImpCabec_MULTPRO2_Fonte12 pnumFileDados, pnumFile, pobjstcPed, pobjstcPedLoc, 0, 0, pLinha
    Else
        mfImpCabec_MULTCOTA_Fonte12 pnumFileDados, pnumFile, pobjstcPed, pobjstcPedLoc, 0, 0, pLinha
    End If


End Sub

Private Sub mpGeraFiguraArqWordSeparado(pArquivo As String, ByRef pRetorno As Byte)
    Dim lAplicacao As Object
    Dim lDocumento As Object
    Dim lArquivo As String
    Dim lArray As Variant
    Dim lI As Integer
    Dim lListaArquivos As String

    On Error GoTo ControleErro
    lArquivo = Dir(gM46V999.gAppPath & Replace(pArquivo, ".WMF", "*.WMF"))
    Do While lArquivo <> ""
        lListaArquivos = lListaArquivos & lArquivo & ";"
        lArquivo = Dir()
    Loop
    lArray = Split(lListaArquivos, ";")
    For lI = (UBound(lArray) - 1) To 0 Step -1
        pRetorno = 0
        Set lAplicacao = CreateObject("Word.Application")

        lAplicacao.Documents.Add

        Set lDocumento = lAplicacao.Documents(1)

        lDocumento.PageSetup.LineNumbering.Active = False
        lDocumento.PageSetup.TopMargin = lAplicacao.CentimetersToPoints(1)
        lDocumento.PageSetup.BottomMargin = lAplicacao.CentimetersToPoints(1)
        lDocumento.PageSetup.LeftMargin = lAplicacao.CentimetersToPoints(1)
        lDocumento.PageSetup.RightMargin = lAplicacao.CentimetersToPoints(1)


        lDocumento.PageSetup.Orientation = 1

        lDocumento.InlineShapes.AddPicture filename:=gM46V999.gAppPath & lArray(lI), _
                                           LinkToFile:=False, SaveWithDocument:=True


        lDocumento.Protect password:="yasuda21", NoReset:=False, Type:=2
        If Dir(gM46V999.gAppPath & Replace(lArray(lI), "WMF", "DOC")) <> "" Then
            Kill gM46V999.gAppPath & Replace(lArray(lI), "WMF", "DOC")
        End If
        lDocumento.SaveAs gM46V999.gAppPath & Replace(lArray(lI), "WMF", "DOC")
        lAplicacao.Documents.Close
        Set lAplicacao = Nothing
        Set lDocumento = Nothing

    Next lI
    pRetorno = 1
ControleErro:
    On Error Resume Next
    Set lAplicacao = Nothing

End Sub

Public Sub gpGeraFiguraArqWordOld(pArquivo As String, ByRef pRetorno As Byte)
    Dim lAplicacao As Object
    Dim lDocumento As Object
    Dim lArquivo As String
    Dim lArray As Variant
    Dim lI As Integer
    Dim lListaArquivos As String

    On Error GoTo ControleErro
    lArquivo = Dir(gM46V999.gAppPath & Replace(pArquivo, ".WMF", "*.WMF"))
    Do While lArquivo <> ""
        lListaArquivos = lListaArquivos & lArquivo & ";"
        lArquivo = Dir()
    Loop
    lArray = Split(lListaArquivos, ";")
    If UBound(lArray) > 0 Then
        pRetorno = 0
        Set lAplicacao = CreateObject("Word.Application")

        lAplicacao.Documents.Add
        lAplicacao.Visible = True

        Set lDocumento = lAplicacao.Documents(1)

        lDocumento.PageSetup.LineNumbering.Active = False

        lDocumento.PageSetup.TopMargin = lAplicacao.CentimetersToPoints(0.5)
        lDocumento.PageSetup.BottomMargin = lAplicacao.CentimetersToPoints(0.5)
        lDocumento.PageSetup.LeftMargin = lAplicacao.CentimetersToPoints(0.5)
        lDocumento.PageSetup.RightMargin = lAplicacao.CentimetersToPoints(0.5)

        'call gM46V999.gpGraLog(0,UBound(lArray) & " folhas.")
        For lI = (UBound(lArray) - 1) To 0 Step -1
            If UCase(lArray(lI)) Like "*ITENS*" Then
                lDocumento.PageSetup.Orientation = 1
            End If

            lDocumento.InlineShapes.AddPicture filename:=gM46V999.gAppPath & lArray(lI), _
                                               LinkToFile:=False, SaveWithDocument:=True
        Next lI

        lDocumento.Protect password:="yasuda21", NoReset:=False, Type:=2
        If Dir(gM46V999.gAppPath & Replace(pArquivo, "WMF", "DOC")) <> "" Then
            Kill gM46V999.gAppPath & Replace(pArquivo, "WMF", "DOC")
        End If


        lDocumento.SaveAs gM46V999.gAppPath & Replace(pArquivo, "WMF", "DOC")
        lAplicacao.WindowState = 1
        Set lAplicacao = Nothing
        Set lDocumento = Nothing

        For lI = (UBound(lArray) - 1) To 0 Step -1
            Kill gM46V999.gAppPath & lArray(lI)
        Next lI

    End If
    pRetorno = 1
ControleErro:
    On Error Resume Next
    Set lAplicacao = Nothing

End Sub

Private Function mfCriDados_MULTLMIE(ByVal pNumArqDados As Integer, ByVal pobjstcPed As stcA46V702B, _
                                     ByRef pEndConfig As String) As Boolean

'Parâmetro de entrada...pNumArqDados...Número do arquivo de dados.

    Dim lAuxiliar2 As String    'Auxiliar - utilizado apenas como exemplo.
    Dim lI As Integer             'Utilizado com For...Next.
    Dim lObjstcCorr As stcA46V705B    'Objeto stcA46V705B  (corretor).
    Dim lobjstcCob As stcA46V704B    'Objeto da classe (stc) de cobertura.
    Dim lValorComis As Single
    Dim lMensagem As String
    Dim vCobert() As String
    Dim lobjstcPedLoc As stcA46V708B    'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
    Dim lobjstcPedCobert As stcA46V704B    'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
    Dim lobjstcClau As stcA46V703B    'Objeto da classe clsA46V703A (TAB_PED_CLAU).
    Dim lobjstcTextoEspec As stcA46V726B
    Dim lQtdCobert As Integer
    Dim lAchou As Boolean
    Dim lPrazo As Integer
    Dim lConcessionaria As Boolean

    On Error GoTo ControleErro

    Dim lPosY As Double
    Dim lNumArqCXY As Integer
    Dim lLinCXY As Long
    Dim lDescFranquiasAux() As String
    Dim lDescFranquias() As String
    Dim lLinhaFranq As Long

    'ROTINA CRIA CXY
    lCountPage = 1
    If gfIsMetaframe() Then
        mEndConfig = "L:\SYAS.CXY"
        mEndArq = "L:\SYAS.DDD"
        Call gM46V999.gpGraLog(0, "Eh Metaframe forcou o L " & mEndConfig)
    Else
        mEndConfig = gM46V999.gAppPath & "SYAS.CXY"
        mEndArq = gM46V999.gAppPath & "SYAS.DDD"
    End If

    lListaArq = mEndArq & "|" & mEndConfig & "|" & gM46V999.gEndImagESPLMI & ";"

    mfCriDados_MULTLMIE = False
    PodeGerarCYX = True
    lNumArqCXY = FreeFile

    On Error Resume Next
    Open mEndConfig For Output Access Write Lock Read Write As #lNumArqCXY

    mfCriDados_Cabec_ESP_LMI pobjstcPed, lLinCXY, lPosY, pNumArqDados, lNumArqCXY, True
    lPosY = 7
    '012. Coberturas (Cabec).
    Print #pNumArqDados, "                          COBERTURAS"
    GeraCYX_LMI lLinCXY, " Coberturas Cabec", lPosY, 0.5, False, "S", lNumArqCXY
    Print #pNumArqDados, "    LMI ÚNICO (R$)"
    GeraCYX_LMI lLinCXY, " Coberturas Cabec", lPosY, 17.3, True, "S", lNumArqCXY

    lQtdCobert = 0
    ReDim lDescFranquias(50)
    lLinhaFranq = 0
    lConcessionaria = False
    'Carrega um array com o maior valor de is de cada cobertura
    For Each lobjstcPedLoc In pobjstcPed.ColecaoPedLoc
        If val(lobjstcPedLoc.Cod_Clasf) = gM46V999.gfBuscarDePara(22, 2) And val(lobjstcPedLoc.Cod_Compl_Clasf) = gM46V999.gfBuscarDePara(22, 2, 13) Then
            lConcessionaria = True
        End If
        For Each lobjstcPedCobert In lobjstcPedLoc.ColecaoPedCobert
            lAchou = False
            For lI = 1 To lQtdCobert
                If vCobert(0, lI) = lobjstcPedCobert.Cod_Cobert Then
                    lAchou = True
                    Exit For
                End If
            Next
            If Not lAchou Then
                lQtdCobert = lQtdCobert + 1
                ReDim Preserve vCobert(5, lQtdCobert)
                vCobert(0, lQtdCobert) = lobjstcPedCobert.Cod_Cobert
                vCobert(1, lQtdCobert) = lobjstcPedCobert.gNomeCobertura
                vCobert(2, lQtdCobert) = lobjstcPedCobert.Val_IS
                vCobert(3, lQtdCobert) = lobjstcPedCobert.Val_Franq_Min
                vCobert(4, lQtdCobert) = lobjstcPedCobert.Num_Meses_PI
                vCobert(5, lQtdCobert) = lobjstcPedCobert.Per_Franq

                'Call mpDescricaoFranquia_CorretorCobertura(lObjstcPedLoc, pobjstcPed, lobjstcPedCobert.Cod_Cobert, lobjstcPedCobert.Val_Franq_Min, lDescFranquiasAux, 50)

                ''Não deve mais buscar pois já está gravado
                'Call gM46V999.mpNovaBuscaDescricaoFranquia(pobjstcPed, lObjstcPedLoc, 50, lobjstcPedCobert.Cod_Cobert, lDescFranquiasAux)

                For lI = 0 To 50
                    If lDescFranquiasAux(lI) = "" Then Exit For
                    lDescFranquias(lLinhaFranq) = lDescFranquiasAux(lI)
                    lLinhaFranq = lLinhaFranq + 1
                Next lI
                If lConcessionaria = True Then
                    '''                    Select Case lobjstcPedCobert.Cod_Cobert
                    '''                        Case 56
                    '''                            lDescFranquias(lLinhaFranq) = "Cobertura 56 - Pátios - até 100 Km: Nos sinistros ocorridos com veículos embarcados em guinchos/cegonhas de terceiros,"
                    '''                            lLinhaFranq = lLinhaFranq + 1
                    '''                            lDescFranquias(lLinhaFranq) = "a Franquia fica substituída pela participação obrigatória do segurado em 10% prej. mín. R$ 1.500,00/R$ 3.000,00 - "
                    '''                            lLinhaFranq = lLinhaFranq + 1
                    '''                            lDescFranquias(lLinhaFranq) = "Nacionais/Importados."
                    '''                            lLinhaFranq = lLinhaFranq + 1
                    '''                        Case 57
                    '''                            lDescFranquias(lLinhaFranq) = "Cobertura 57 - R C Concessionárias - até 100 Km: Nos sinistros ocorridos com veículos embarcados em guinchos/cegonhas "
                    '''                            lLinhaFranq = lLinhaFranq + 1
                    '''                            lDescFranquias(lLinhaFranq) = "de terceiros, a Franquia fica substituída pela participação obrigatória do segurado em 10% prej. mín. "
                    '''                            lLinhaFranq = lLinhaFranq + 1
                    '''                            lDescFranquias(lLinhaFranq) = "R$ 1.500,00/R$ 3.000,00 - Nacionais/Importados."
                    '''                            lLinhaFranq = lLinhaFranq + 1
                    '''                        Case 122
                    '''                            lDescFranquias(lLinhaFranq) = "Cobertura 122 - Pátios - até 200 Km: Nos sinistros ocorridos com veículos embarcados em guinchos/cegonhas de terceiros, "
                    '''                            lLinhaFranq = lLinhaFranq + 1
                    '''                            lDescFranquias(lLinhaFranq) = "a Franquia fica substituída pela participação obrigatória do segurado em 10% prej. mín. R$ 1.500,00/R$ 3.000,00 - "
                    '''                            lLinhaFranq = lLinhaFranq + 1
                    '''                            lDescFranquias(lLinhaFranq) = "Nacionais/Importados."
                    '''                            lLinhaFranq = lLinhaFranq + 1
                    '''                        Case 123
                    '''                            lDescFranquias(lLinhaFranq) = "Cobertura 123 - Pátios - até 300 Km : Nos sinistros ocorridos com veículos embarcados em guinchos/cegonhas de terceiros, "
                    '''                            lLinhaFranq = lLinhaFranq + 1
                    '''                            lDescFranquias(lLinhaFranq) = "a Franquia fica substituída pela participação obrigatória do segurado em 10% prej. mín. R$ 1.500,00/R$ 3.000,00 - "
                    '''                            lLinhaFranq = lLinhaFranq + 1
                    '''                            lDescFranquias(lLinhaFranq) = "Nacionais/Importados."
                    '''                            lLinhaFranq = lLinhaFranq + 1
                    '''                        Case 124
                    '''                            lDescFranquias(lLinhaFranq) = "Cobertura 124 - R C Concessionárias - até 200 Km:  Nos sinistros ocorridos com veículos embarcados em guinchos/cegonhas "
                    '''                            lLinhaFranq = lLinhaFranq + 1
                    '''                            lDescFranquias(lLinhaFranq) = "de terceiros, a Franquia fica substituída pela participação obrigatória do segurado em 10% prej. mín. "
                    '''                            lLinhaFranq = lLinhaFranq + 1
                    '''                            lDescFranquias(lLinhaFranq) = "R$ 2.000,00/R$ 3.000,00 - Nacionais/Importados."
                    '''                            lLinhaFranq = lLinhaFranq + 1
                    '''                        Case 125
                    '''                            lDescFranquias(lLinhaFranq) = "Cobertura 125 - R C Concessionárias - até 300 Km:  Nos sinistros ocorridos com veículos embarcados em guinchos/cegonhas "
                    '''                            lLinhaFranq = lLinhaFranq + 1
                    '''                            lDescFranquias(lLinhaFranq) = "de terceiros, a Franquia fica substituída pela participação obrigatória do segurado em 10% prej. mín. "
                    '''                            lLinhaFranq = lLinhaFranq + 1
                    '''                            lDescFranquias(lLinhaFranq) = "R$ 2.000,00/R$ 3.000,00 - Nacionais/Importados."
                    '''                            lLinhaFranq = lLinhaFranq + 1
                    '''                    End Select
                End If
            Else
                If vCobert(2, lI) < lobjstcPedCobert.Val_IS Then
                    vCobert(2, lI) = lobjstcPedCobert.Val_IS
                End If
                If vCobert(3, lI) < lobjstcPedCobert.Val_Franq_Min Then
                    vCobert(3, lI) = lobjstcPedCobert.Val_Franq_Min
                End If
                If vCobert(4, lI) < lobjstcPedCobert.Num_Meses_PI Then
                    vCobert(4, lI) = lobjstcPedCobert.Num_Meses_PI
                End If
                If vCobert(5, lI) < lobjstcPedCobert.Per_Franq Then
                    vCobert(5, lI) = lobjstcPedCobert.Per_Franq
                End If
            End If
        Next
    Next

    For lI = 1 To lQtdCobert
        lAuxiliar2 = Format(vCobert(0, lI), "0000")    'código da cobertura
        GeraCYX_LMI lLinCXY, "Descricao da cobertura", lPosY, 0.5, False, "N", lNumArqCXY
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        lAuxiliar2 = vCobert(1, lI)  'Descrição da cobertura
        GeraCYX_LMI lLinCXY, "Descricao da cobertura", lPosY, 1.5, False, "N", lNumArqCXY
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        lAuxiliar2 = Format(vCobert(2, lI), "standard")  'Maior Valor do LMI da cobertura
        lAuxiliar2 = gfPreencheBranco(lAuxiliar2, 20)
        GeraCYX_LMI lLinCXY, "Valor do LMI da cobertura", lPosY, 17.3, True, "N", lNumArqCXY
        Print #pNumArqDados, lAuxiliar2
    Next

    GeraCYX_LMI lLinCXY, "Linha Vazia", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed

    '020 Prêmio  .

    'Soma da Comissão para saber se o prêmio é NET
    If frmT46V101A.objstcPedido.ColecaoPedCorr.Count > 0 Then
        Set lObjstcCorr = frmT46V101A.objstcPedido.ColecaoPedCorr.Item(1)
        lValorComis = lObjstcCorr.Per_Corr
        If frmT46V101A.objstcPedido.ColecaoPedCorr.Count > 1 Then
            Set lObjstcCorr = frmT46V101A.objstcPedido.ColecaoPedCorr.Item(2)
            lValorComis = lValorComis + lObjstcCorr.Per_Corr
            If frmT46V101A.objstcPedido.ColecaoPedCorr.Count > 2 Then
                Set lObjstcCorr = frmT46V101A.objstcPedido.ColecaoPedCorr.Item(3)
                lValorComis = lValorComis + lObjstcCorr.Per_Corr
            End If
        End If
    End If

    GeraCYX_LMI lLinCXY, "Premio", lPosY, 0.5, False, "S", lNumArqCXY, pNumArqDados, pobjstcPed
    If lValorComis = 0 Then
        Print #pNumArqDados, "Prêmio Net: R$"
    Else
        Print #pNumArqDados, "Prêmio Líquido: R$"
    End If


    lAuxiliar2 = Format$(frmT46V101A.objstcPedido.VAL_PR, "standard")
    lAuxiliar2 = gfPreencheBranco(lAuxiliar2, 20)
    GeraCYX_LMI lLinCXY, "Premio", lPosY, 4.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)


    'Controle de página
    lNumPagina = 1
    lTotPagina = 1

    If lQtdCobert >= 13 Then
        lTotPagina = lTotPagina + Int(pobjstcPed.ColecaoPedLoc.Count / 2)
        If pobjstcPed.ColecaoPedLoc.Count Mod 2 > 0 Then
            lTotPagina = lTotPagina + 1
        End If
    ElseIf lQtdCobert >= 6 Then
        lTotPagina = lTotPagina + Int(pobjstcPed.ColecaoPedLoc.Count / 4)
        If pobjstcPed.ColecaoPedLoc.Count Mod 4 > 0 Then
            lTotPagina = lTotPagina + 1
        End If
    Else
        lTotPagina = lTotPagina + Int(pobjstcPed.ColecaoPedLoc.Count / 6)
        If pobjstcPed.ColecaoPedLoc.Count Mod 6 > 0 Then
            lTotPagina = lTotPagina + 1
        End If
    End If

    GeraCYX_LMI lLinCXY, "Linha Vazia", lPosY, 0.5, True, "N", lNumArqCXY

    '013. Label Riscos Cobertos.
    GeraCYX_LMI lLinCXY, "Label Riscos Cobertos", lPosY, 0.5, False, "S", lNumArqCXY
    Print #pNumArqDados, "Riscos Cobertos:"


    '014. Descrição Riscos Cobertos.
    lAuxiliar2 = " "
    Select Case pobjstcPed.Cod_Ramo
    Case e_CodRamo_Condominio
        lAuxiliar2 = "Conforme Condições Gerais do Produto Yasuda Condomínio"
    Case e_CodRamo_Empresarial
        lAuxiliar2 = "Conforme Condições Gerais do Produto Yasuda Empresarial"
    Case e_CodRamo_Residencial
        lAuxiliar2 = "Conforme Condições Gerais do Produto Yasuda Residencial"
    End Select
    GeraCYX_LMI lLinCXY, "Label Riscos Cobertos", lPosY, 4.5, True, "N", lNumArqCXY
    Print #pNumArqDados, lAuxiliar2


    GeraCYX_LMI lLinCXY, "Linha Vazia", lPosY, 0.5, True, "N", lNumArqCXY

    '015. Resumo: quantidade de itens.
    GeraCYX_LMI lLinCXY, "Locais de Risco", lPosY, 0.5, False, "S", lNumArqCXY
    Print #pNumArqDados, "Locais de Risco:"



    GeraCYX_LMI lLinCXY, "Locais de Risco", lPosY, 4.5, True, "N", lNumArqCXY
    Print #pNumArqDados, "(" & pobjstcPed.ColecaoPedLoc.Count & " locais) conforme especificação anexa"

    GeraCYX_LMI lLinCXY, "Linha Vazia", lPosY, 0.5, True, "N", lNumArqCXY

    '016. Valor em Risco.

    GeraCYX_LMI lLinCXY, "Valor em Risco", lPosY, 0.5, False, "S", lNumArqCXY
    Print #pNumArqDados, "Valor em Risco:"


    GeraCYX_LMI lLinCXY, "Valor em Risco", lPosY, 4.5, True, "N", lNumArqCXY
    Print #pNumArqDados, " Conforme especificação anexa"

    GeraCYX_LMI lLinCXY, "Linha Vazia", lPosY, 0.5, True, "N", lNumArqCXY

    '017 Ocupação.
    Print #pNumArqDados, "Ocupação:"
    GeraCYX_LMI lLinCXY, "Ocupacao", lPosY, 0.5, False, "S", lNumArqCXY, pNumArqDados, pobjstcPed


    Print #pNumArqDados, " Conforme especificação anexa"
    GeraCYX_LMI lLinCXY, "Ocupacao", lPosY, 4.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed



    GeraCYX_LMI lLinCXY, "Linha Vazia", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed

    '018 Franquia.
    If lPosY > 26.4 Then
        GeraCYX_LMI lLinCXY, "Linha Vazia", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
    End If

    Print #pNumArqDados, "Franquias/Período indenitário:"
    GeraCYX_LMI lLinCXY, "Franquia", lPosY, 0.5, True, "S", lNumArqCXY, pNumArqDados, pobjstcPed


    For lI = 1 To lQtdCobert
        If vCobert(3, lI) > 0 Then  'Tem franquia
            lAuxiliar2 = vCobert(1, lI) & ": "  'Descrição da cobertura
            Print #pNumArqDados, lAuxiliar2
            GeraCYX_LMI lLinCXY, "Franquia", lPosY, 0.5, False, "N", lNumArqCXY, pNumArqDados, pobjstcPed

            lAuxiliar2 = ""
            If vCobert(5, lI) > 0 Then
                lAuxiliar2 = vCobert(5, lI) & "% dos prej.indeniz., min: "
            End If
            lAuxiliar2 = lAuxiliar2 & "R$ " & Format$(vCobert(3, lI), "standard")
            Print #pNumArqDados, lAuxiliar2
            GeraCYX_LMI lLinCXY, "Franquia", lPosY, 13, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
        ElseIf vCobert(4, lI) > 0 Then  'Pi indenitário
            lAuxiliar2 = vCobert(1, lI) & ": "    'Descrição da cobertura
            Print #pNumArqDados, lAuxiliar2
            GeraCYX_LMI lLinCXY, "Franquia", lPosY, 0.5, False, "N", lNumArqCXY, pNumArqDados, pobjstcPed

            lAuxiliar2 = "PI = " & vCobert(4, lI) & " meses"
            Call mpPrazoTabCoefPI(pobjstcPed, CInt(vCobert(0, lI)), CInt(vCobert(4, lI)), lPrazo)
            If lPrazo <> 0 Then
                lAuxiliar2 = lAuxiliar2 & ": " & lPrazo & " dias"
            End If
            Print #pNumArqDados, lAuxiliar2
            GeraCYX_LMI lLinCXY, "Franquia", lPosY, 13, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
        End If
    Next

    For lI = 0 To 50
        If lDescFranquias(lI) = "" Then Exit For
        lAuxiliar2 = lDescFranquias(lI)
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        GeraCYX_LMI lLinCXY, "Franquia", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
    Next lI

    GeraCYX_LMI lLinCXY, "Linha Vazia", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed

    '019 Cláusulas e Condições.
    Print #pNumArqDados, "Cláusulas e Condições:"
    GeraCYX_LMI lLinCXY, "Clausulas e Condicoes:", lPosY, 0.5, False, "S", lNumArqCXY, pNumArqDados, pobjstcPed

    Print #pNumArqDados, "Ratificam-se as seguintes Cláusulas e Condições:"
    GeraCYX_LMI lLinCXY, "Clausulas e Condicoes", lPosY, 4.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed

    lAuxiliar2 = " "
    Select Case pobjstcPed.Cod_Ramo
    Case e_CodRamo_Condominio
        If pobjstcPed.Cod_Produto >= 1080 Then
            lAuxiliar2 = "a) Condições Gerais do Produto Condomínio (versão 10/2014)"
        Else
            lAuxiliar2 = "a) Condições Gerais do Produto Yasuda Condomínio (versão 09/2008)"
        End If
            
        Print #pNumArqDados, lAuxiliar2
        GeraCYX_LMI lLinCXY, "Clausulas e Condicoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
        lI = 98
    Case e_CodRamo_Empresarial
        If pobjstcPed.Cod_Produto >= 1080 Then
            lAuxiliar2 = "a) Condições Gerais do Produto Empresarial (versão 10/2014)"
        Else
            lAuxiliar2 = "a) Condições Gerais do Produto Yasuda Empresarial (versão 06/2009)"
        End If
            
        Print #pNumArqDados, lAuxiliar2
        GeraCYX_LMI lLinCXY, "Clausulas e Condicoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
        '1200029 - Clausula 315 e Texto LMI
        '''            lAuxiliar2 = "b) Cláusula 315- LMI Único "
        '''            Print #pNumArqDados, lAuxiliar2
        '''            GeraCYX_LMI lLinCXY, "Clausulas e Condicoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
        lI = 98
    Case e_CodRamo_Residencial
        If pobjstcPed.Cod_Produto >= 1080 Then
            lAuxiliar2 = "a) Condições Gerais do Produto Residencial (versão 10/2014)"
        Else
            lAuxiliar2 = "a) Condições Gerais do Produto Yasuda Residencial (versão 09/2008)"
        End If
            
        Print #pNumArqDados, lAuxiliar2
        GeraCYX_LMI lLinCXY, "Clausulas e Condicoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
        lI = 98
    End Select
    Dim lClausula As String
    lClausula = ""
    For Each lobjstcPedLoc In pobjstcPed.ColecaoPedLoc
        For Each lobjstcClau In lobjstcPedLoc.ColecaoPedClau
            If lobjstcClau.Cod_Clau >= 310 Or (lobjstcClau.Cod_Ramo = e_CodRamo_Residencial And lobjstcClau.Cod_Clau > 300) Or pobjstcPed.LMI_Unico = 1 Then
                If Not lClausula Like "*;" & lobjstcClau.Cod_Clau & ";*" Then
                    lAuxiliar2 = Chr(lI) & ") Cláusula " & lobjstcClau.Cod_Clau & " - " & lobjstcClau.gDescClausula
                    If lobjstcClau.gDescClausula Like "*CLAUSULA BENEFICIARIA*" Then
                        lAuxiliar2 = Trim(lAuxiliar2) & "(*)"
                    End If
                    Print #pNumArqDados, lAuxiliar2
                    GeraCYX_LMI lLinCXY, "Clausulas e Condicoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
                    lI = lI + 1
                    lClausula = lClausula & ";" & lobjstcClau.Cod_Clau & ";"
                End If
            End If
        Next lobjstcClau
    Next

    If pobjstcPed.Cod_Ramo = e_CodRamo_Residencial Or pobjstcPed.Cod_Ramo = e_CodRamo_Condominio Then
        lAuxiliar2 = Chr(lI) & ") CLÁUSULA PARTICULAR L.M.I. Único - Declara-se para os devidos fins e efeitos que, este seguro foi contratado sob "
        lI = lI + 1
        Print #pNumArqDados, lAuxiliar2
        GeraCYX_LMI lLinCXY, "Clausulas e Condicoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
        lAuxiliar2 = "a forma de LMI Único para todos os locais segurados especificados na presente apólice. Ao contrário do que possa constar "
        Print #pNumArqDados, lAuxiliar2
        GeraCYX_LMI lLinCXY, "Clausulas e Condicoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
        lAuxiliar2 = "nas condições gerais do presente contrato de seguro, fica entendido e acordado que em caso de eventual sinistro, a "
        Print #pNumArqDados, lAuxiliar2
        GeraCYX_LMI lLinCXY, "Clausulas e Condicoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
        lAuxiliar2 = "indenização não poderá exceder ao LMI de cada cobertura contratada, limitado ao VR declarado para cada local."
        Print #pNumArqDados, lAuxiliar2
        GeraCYX_LMI lLinCXY, "Clausulas e Condicoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
    End If


    GeraCYX_LMI lLinCXY, "Linha Vazia", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed

    '021 Observações .

    Print #pNumArqDados, "Observações:"
    GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "S", lNumArqCXY, pNumArqDados, pobjstcPed




    If pobjstcPed.Cod_Ramo = e_CodRamo_Empresarial And lConcessionaria = True Then
        Select Case pobjstcPed.objstcCtrlEmis.Cod_Corr
        Case 8001, 8002, 8003, 8004, 8005, 8006, 8008, 8011, 8014, 8028
            lAuxiliar2 = "Mercadorias e matérias-primas de terceiros em poder do segurado: seus valores deverão ser considerados na composição "
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed

            lAuxiliar2 = "do Valor Risco Total Declarado, devendo ser observado os bens excluídos das Condições Gerais do Produto Multi"
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed

            lAuxiliar2 = "Seguro Empresarial;"
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
        End Select
    End If

    'Busco os textos salvos na especificação antiga.

    If pobjstcPed.ColecaoTextoEspec.Count > 0 Then
        For Each lobjstcTextoEspec In pobjstcPed.ColecaoTextoEspec
            If InStr(1, lobjstcTextoEspec.Espec_lin1, "** FRANQUIA **", vbTextCompare) > 0 Or _
               InStr(1, lobjstcTextoEspec.Espec_lin2, "** FRANQUIA **", vbTextCompare) > 0 Then
                Exit For
            End If
            lAuxiliar2 = lobjstcTextoEspec.Espec_lin1
            If lAuxiliar2 Like "*CLAUSULA BENEFICIARIA*" Then
                lAuxiliar2 = "(*)" & lAuxiliar2
            End If
            Print #pNumArqDados, lAuxiliar2
            GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
            If gM46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin2) Then
                Print #pNumArqDados, lobjstcTextoEspec.Espec_lin2
                GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
            End If
            If gM46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin3) Then
                Print #pNumArqDados, lobjstcTextoEspec.Espec_lin3
                GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
            End If
            If gM46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin4) Then
                Print #pNumArqDados, lobjstcTextoEspec.Espec_lin4
                GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
            End If
            If gM46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin5) Then
                Print #pNumArqDados, lobjstcTextoEspec.Espec_lin5
                GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
            End If
            If gM46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin6) Then
                Print #pNumArqDados, lobjstcTextoEspec.Espec_lin6
                GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
            End If
            If gM46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin7) Then
                Print #pNumArqDados, lobjstcTextoEspec.Espec_lin7
                GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
            End If
            If gM46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin8) Then
                Print #pNumArqDados, lobjstcTextoEspec.Espec_lin8
                GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
            End If
            If gM46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin9) Then
                Print #pNumArqDados, lobjstcTextoEspec.Espec_lin9
                GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
            End If
            If gM46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin10) Then
                Print #pNumArqDados, lobjstcTextoEspec.Espec_lin10
                GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
            End If
            If gM46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin11) Then
                Print #pNumArqDados, lobjstcTextoEspec.Espec_lin11
                GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
            End If
            If gM46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin12) Then
                Print #pNumArqDados, lobjstcTextoEspec.Espec_lin12
                GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
            End If
            If gM46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin13) Then
                Print #pNumArqDados, lobjstcTextoEspec.Espec_lin13
                GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
            End If
            If gM46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin14) Then
                Print #pNumArqDados, lobjstcTextoEspec.Espec_lin14
                GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
            End If
            If gM46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin15) Then
                Print #pNumArqDados, lobjstcTextoEspec.Espec_lin15
                GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
            End If
            If gM46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin16) Then
                Print #pNumArqDados, lobjstcTextoEspec.Espec_lin16
                GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
            End If
            If gM46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin17) Then
                Print #pNumArqDados, lobjstcTextoEspec.Espec_lin17
                GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
            End If
            If gM46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin18) Then
                Print #pNumArqDados, lobjstcTextoEspec.Espec_lin18
                GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
            End If
            If gM46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin19) Then
                Print #pNumArqDados, lobjstcTextoEspec.Espec_lin19
                GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
            End If
            If gM46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin20) Then
                Print #pNumArqDados, lobjstcTextoEspec.Espec_lin20
                GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
            End If
            Print #pNumArqDados, " "
            GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
        Next lobjstcTextoEspec
    End If


    'Rodape

    Print #pNumArqDados, "Página: " & lNumPagina & "/@@TOTALPAG@@"
    GeraCYX_LMI lLinCXY, "Pagina", 28.5, 18, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
    lNumPagina = lNumPagina + 1

    Print #lNumArqCXY, " "
    Close #lNumArqCXY

    'Print #pNumArqDados, pobjstcPed.Cod_User_Liber & " - " & gM46V999.gfForData(2, pobjstcPed.Dat_Liber) & " - " & Mid(Format(pobjstcPed.Hor_Liber, "0000"), 1, 2) & ":" & Mid(Format(pobjstcPed.Hor_Liber, "0000"), 3, 2)

    mfCriDados_MULTLMIE = True

    Exit Function
ControleErro:
    Call gM46V999.gpGraLog(1, "mfCriDados_MULTLMIE ERRO - " & Err & " - " & Error)

End Function

Private Function mfSplitEX(ByVal pTexto As String, ByVal pDelimitador As Byte, ByVal pMaximo As Integer) As String()
    Dim Retorno() As String
    Dim i As Integer
    Dim pos As Integer
    
    ReDim Retorno(1000)

    i = 0
    Do While Len(pTexto) > 0

        If InStr(pTexto, Chr(pDelimitador)) <= pMaximo And InStr(pTexto, Chr(pDelimitador)) > 0 Then
            Retorno(i) = Left(pTexto, InStr(pTexto, Chr(pDelimitador)))
            pTexto = Mid(pTexto, Len(Retorno(i)) + 2, Len(pTexto))
        Else
            If InStr(pTexto, Chr(pDelimitador)) = 0 Then
                pos = InStrRev(Left(pTexto, pMaximo), " ")
                If pos = 0 Then
                    pos = pMaximo
                End If
                Retorno(i) = Left(pTexto, pos)
                pTexto = Mid(pTexto, Len(Retorno(i)) + 1, Len(pTexto))
            Else
                Retorno(i) = Left(pTexto, pMaximo)
                pTexto = Mid(pTexto, Len(Retorno(i)) + 1, Len(pTexto))
            End If
        End If

        i = i + 1
    Loop
    
    If i > 1 Then
        ReDim Preserve Retorno(i - 1)
    Else
        ReDim Preserve Retorno(1)
    End If
    
    mfSplitEX = Retorno
End Function

Private Function mfCriDados_COT_LMI(ByVal pNumArqDados As Integer, ByVal pobjstcPed As stcA46V702B, _
                                    ByRef pEndConfig As String) As Boolean

'Função: cria arquivo de dados - cotação coletiva - capa.

'Parâmetro de entrada...pNumArqDados...Número do arquivo de dados.

    Dim lAuxiliar2 As String    'Auxiliar - utilizado apenas como exemplo.
    Dim lI As Integer             'Utilizado com For...Next.
    Dim lObjstcCorr As stcA46V705B    'Objeto stcA46V705B  (corretor).
    Dim lobjstcCob As stcA46V704B    'Objeto da classe (stc) de cobertura.
    Dim lobjstcCobLMI As stcA46V732B    'Objeto da classe (stc) de LMI
    Dim lValorComis As Single
    Dim lMensagem As String
    Dim vCobert() As String
    Dim lobjstcPedLoc As stcA46V708B    'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
    Dim lObjstcPedLocTmp As stcA46V708B
    Dim lobjstcPedCobert As stcA46V704B    'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
    Dim lobjstcClau As stcA46V703B    'Objeto da classe clsA46V703A (TAB_PED_CLAU).
    Dim lobjstcPedDizer As stcA46V722B
    Dim lQtdCobert As Integer
    Dim lAchou As Boolean
    Dim lPrazo As Integer
    Dim lConcessionaria As Boolean
    Dim lTamanho As Integer
    Dim lValorParc As Variant
    Dim lOrdem As Integer
    Dim lAuxiliar1 As Double    'Auxiliar - utilizado apenas como exemplo.
    Dim lPriParcela As String    'Primeira parcela.
    Dim lDemParcelas As String    'Demais parcelas.
    Dim lobjstcJur As stcA46V711B
    Dim lJuros As String    'Juros.
    Dim lMsgInspecao As String
    Dim lTextoCong As String
    Dim lTemDesc As Boolean
    Dim blnCorretorVolks As Boolean
    Dim lTemAssistencia As Boolean
    Dim lContador As Integer
    On Error GoTo ControleErro

    Dim lPosY As Double
    Dim lNumArqCXY As Integer
    Dim lLinCXY As Long


    Dim lDescFranquiasAux() As String
    Dim lDescFranquias() As String
    Dim lLinhaFranq As Long

    Dim lObs As Variant
    Dim lCount As Integer
    Dim lPos As Long
    Dim lDescImportante() As String
    Dim lTam As Integer
    Dim Aux As Integer
    Dim lListClausula As String
    Dim lTemObs As Boolean
    Dim TemInspecaoRisco As Boolean
    Dim Titulo As String


    'ROTINA CRIA CXY
    lCountPage = 1
    If gfIsMetaframe() Then
        mEndConfig = "L:\SYAS.CXY"
        mEndArq = "L:\SYAS.DDD"
        Call gM46V999.gpGraLog(0, "Eh Metaframe forcou o L " & mEndConfig)
    Else
        mEndConfig = gM46V999.gAppPath & "SYAS.CXY"
        mEndArq = gM46V999.gAppPath & "SYAS.DDD"
    End If
    lTemDesc = False
    lListaArq = mEndArq & "|" & mEndConfig & "|" & gM46V999.gEndImagCotLMI & ";"


    mfCriDados_COT_LMI = False
    PodeGerarCYX = True
    lNumArqCXY = FreeFile

    On Error Resume Next
    Open mEndConfig For Output Access Write Lock Read Write As #lNumArqCXY

    mfCriDados_Cabec_LMI pobjstcPed, lLinCXY, lPosY, pNumArqDados, lNumArqCXY

    '012. Coberturas (Cabec).
    Print #pNumArqDados, "                          COBERTURAS"
    GeraCYX_LMI lLinCXY, " Coberturas Cabec", lPosY, 0.5, False, "S", lNumArqCXY
    Print #pNumArqDados, "    LMI ÚNICO (R$)"
    GeraCYX_LMI lLinCXY, " Coberturas Cabec", lPosY, 17.3, True, "S", lNumArqCXY

    ReDim lDescFranquias(50)
    lQtdCobert = 0
    'Carrega um array com o maior valor de is de cada cobertura
    lConcessionaria = False
    lTemAssistencia = False
    For Each lobjstcPedLoc In pobjstcPed.ColecaoPedLoc
        If val(lobjstcPedLoc.Cod_Clasf) = gM46V999.gfBuscarDePara(22, 2) And val(lobjstcPedLoc.Cod_Compl_Clasf) = gM46V999.gfBuscarDePara(22, 2, 13) Then
            lConcessionaria = True
        End If
        For Each lobjstcPedCobert In lobjstcPedLoc.ColecaoPedCobert
            lAchou = False

            If lobjstcPedCobert.Per_Desc_Coml <> 0 Then
                lTemDesc = True
            End If

            For lI = 1 To lQtdCobert
                If vCobert(0, lI) = lobjstcPedCobert.Cod_Cobert Then
                    lAchou = True
                    Exit For
                End If
            Next
            If Not lAchou Then
                lQtdCobert = lQtdCobert + 1
                ReDim Preserve vCobert(6, lQtdCobert)
                vCobert(0, lQtdCobert) = lobjstcPedCobert.Cod_Cobert
                vCobert(1, lQtdCobert) = lobjstcPedCobert.gNomeCobertura
                vCobert(2, lQtdCobert) = lobjstcPedCobert.Val_IS
                vCobert(3, lQtdCobert) = lobjstcPedCobert.Val_Franq_Min
                vCobert(4, lQtdCobert) = lobjstcPedCobert.Num_Meses_PI
                vCobert(5, lQtdCobert) = lobjstcPedCobert.Per_Franq
                vCobert(6, lQtdCobert) = lobjstcPedCobert.TextoFranquia

                'aqui
                'Call mpDescricaoFranquia_CorretorCobertura(lObjstcPedLoc, pobjstcPed, lobjstcPedCobert.Cod_Cobert, lobjstcPedCobert.Val_Franq_Min, lDescFranquiasAux, 50)

                'Não deve mais buscar pois já está gravado
                'Call gM46V999.mpNovaBuscaDescricaoFranquia(pobjstcPed, lObjstcPedLoc, 50, lobjstcPedCobert.Cod_Cobert, lDescFranquiasAux)

                For lI = 0 To 50
                    If lDescFranquiasAux(lI) = "" Then Exit For
                    lDescFranquias(lLinhaFranq) = lDescFranquiasAux(lI)
                    lLinhaFranq = lLinhaFranq + 1
                Next lI

                If lConcessionaria = True Then
                    '''                    Select Case lobjstcPedCobert.Cod_Cobert
                    '''                        Case 56
                    '''                            lDescFranquias(lLinhaFranq) = "Cobertura 56 - Pátios - até 100 Km: Nos sinistros ocorridos com veículos embarcados em guinchos/cegonhas de terceiros,"
                    '''                            lLinhaFranq = lLinhaFranq + 1
                    '''                            lDescFranquias(lLinhaFranq) = "a Franquia fica substituída pela participação obrigatória do segurado em 10% prej. mín. R$ 1.500,00/R$ 3.000,00 - "
                    '''                            lLinhaFranq = lLinhaFranq + 1
                    '''                            lDescFranquias(lLinhaFranq) = "Nacionais/Importados."
                    '''                            lLinhaFranq = lLinhaFranq + 1
                    '''                        Case 57
                    '''                            lDescFranquias(lLinhaFranq) = "Cobertura 57 - R C Concessionárias - até 100 Km: Nos sinistros ocorridos com veículos embarcados em guinchos/cegonhas "
                    '''                            lLinhaFranq = lLinhaFranq + 1
                    '''                            lDescFranquias(lLinhaFranq) = "de terceiros, a Franquia fica substituída pela participação obrigatória do segurado em 10% prej. mín. "
                    '''                            lLinhaFranq = lLinhaFranq + 1
                    '''                            lDescFranquias(lLinhaFranq) = "R$ 1.500,00/R$ 3.000,00 - Nacionais/Importados."
                    '''                            lLinhaFranq = lLinhaFranq + 1
                    '''                        Case 122
                    '''                            lDescFranquias(lLinhaFranq) = "Cobertura 122 - Pátios - até 200 Km: Nos sinistros ocorridos com veículos embarcados em guinchos/cegonhas de terceiros, "
                    '''                            lLinhaFranq = lLinhaFranq + 1
                    '''                            lDescFranquias(lLinhaFranq) = "a Franquia fica substituída pela participação obrigatória do segurado em 10% prej. mín. R$ 1.500,00/R$ 3.000,00 - "
                    '''                            lLinhaFranq = lLinhaFranq + 1
                    '''                            lDescFranquias(lLinhaFranq) = "Nacionais/Importados."
                    '''                            lLinhaFranq = lLinhaFranq + 1
                    '''                        Case 123
                    '''                            lDescFranquias(lLinhaFranq) = "Cobertura 123 - Pátios - até 300 Km : Nos sinistros ocorridos com veículos embarcados em guinchos/cegonhas de terceiros, "
                    '''                            lLinhaFranq = lLinhaFranq + 1
                    '''                            lDescFranquias(lLinhaFranq) = "a Franquia fica substituída pela participação obrigatória do segurado em 10% prej. mín. R$ 1.500,00/R$ 3.000,00 - "
                    '''                            lLinhaFranq = lLinhaFranq + 1
                    '''                            lDescFranquias(lLinhaFranq) = "Nacionais/Importados."
                    '''                            lLinhaFranq = lLinhaFranq + 1
                    '''                        Case 124
                    '''                            lDescFranquias(lLinhaFranq) = "Cobertura 124 - R C Concessionárias - até 200 Km:  Nos sinistros ocorridos com veículos embarcados em guinchos/cegonhas "
                    '''                            lLinhaFranq = lLinhaFranq + 1
                    '''                            lDescFranquias(lLinhaFranq) = "de terceiros, a Franquia fica substituída pela participação obrigatória do segurado em 10% prej. mín. "
                    '''                            lLinhaFranq = lLinhaFranq + 1
                    '''                            lDescFranquias(lLinhaFranq) = "R$ 2.000,00/R$ 3.000,00 - Nacionais/Importados."
                    '''                            lLinhaFranq = lLinhaFranq + 1
                    '''                        Case 125
                    '''                            lDescFranquias(lLinhaFranq) = "Cobertura 125 - R C Concessionárias - até 300 Km:  Nos sinistros ocorridos com veículos embarcados em guinchos/cegonhas "
                    '''                            lLinhaFranq = lLinhaFranq + 1
                    '''                            lDescFranquias(lLinhaFranq) = "de terceiros, a Franquia fica substituída pela participação obrigatória do segurado em 10% prej. mín. "
                    '''                            lLinhaFranq = lLinhaFranq + 1
                    '''                            lDescFranquias(lLinhaFranq) = "R$ 2.000,00/R$ 3.000,00 - Nacionais/Importados."
                    '''                            lLinhaFranq = lLinhaFranq + 1
                    '''                    End Select
                End If
            Else
                If vCobert(2, lI) < lobjstcPedCobert.Val_IS Then
                    vCobert(2, lI) = lobjstcPedCobert.Val_IS
                End If
                If vCobert(3, lI) < lobjstcPedCobert.Val_Franq_Min Then
                    vCobert(3, lI) = lobjstcPedCobert.Val_Franq_Min
                End If
                If vCobert(4, lI) < lobjstcPedCobert.Num_Meses_PI Then
                    vCobert(4, lI) = lobjstcPedCobert.Num_Meses_PI
                End If
                If vCobert(5, lI) < lobjstcPedCobert.Per_Franq Then
                    vCobert(5, lI) = lobjstcPedCobert.Per_Franq
                End If
            End If
        Next
    Next

    For Each lobjstcCobLMI In pobjstcPed.ColecaoLMI
        Set lobjstcCobLMI.M46V999 = gM46V999
        Set lobjstcCobLMI.M46V111 = gM46V111
        lAuxiliar2 = Format(lobjstcCobLMI.Cod_Cobert, "0000")   'código da cobertura
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        GeraCYX_LMI lLinCXY, "Descricao da cobertura", lPosY, 0.5, False, "N", lNumArqCXY

        lAuxiliar2 = lobjstcCobLMI.gNomeCobertura   'Descrição da cobertura
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        GeraCYX_LMI lLinCXY, "Descricao da cobertura", lPosY, 1.5, False, "N", lNumArqCXY

        lAuxiliar2 = Format(lobjstcCobLMI.Val_IS, "standard")   'Maior Valor do LMI da cobertura
        lAuxiliar2 = gfPreencheBranco(lAuxiliar2, 20)
        Print #pNumArqDados, lAuxiliar2
        GeraCYX_LMI lLinCXY, "Valor do LMI da cobertura", lPosY, 17.3, True, "N", lNumArqCXY
    Next lobjstcCobLMI
    'For lI = 1 To lQtdCobert
    '    lAuxiliar2 = Format(vCobert(0, lI), "0000") 'código da cobertura
    '    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    '    GeraCYX_LMI lLinCXY, "Descricao da cobertura", lPosY, 0.5, False, "N", lNumArqCXY
    '
    '    lAuxiliar2 = vCobert(1, lI)  'Descrição da cobertura
    '    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    '    GeraCYX_LMI lLinCXY, "Descricao da cobertura", lPosY, 1.5, False, "N", lNumArqCXY
    '
    '    lAuxiliar2 = Format(vCobert(2, lI), "standard")  'Maior Valor do LMI da cobertura
    '    lAuxiliar2 = gfPreencheBranco(lAuxiliar2, 20)
    '    Print #pNumArqDados, lAuxiliar2
    '    GeraCYX_LMI lLinCXY, "Valor do LMI da cobertura", lPosY, 17.3, True, "N", lNumArqCXY
    'Next

    '020 Prêmio  .

    'Soma da Comissão para saber se o prêmio é NET
    If pobjstcPed.ColecaoPedCorr.Count > 0 Then
        Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(1)
        lValorComis = lObjstcCorr.Per_Corr
        If pobjstcPed.ColecaoPedCorr.Count > 1 Then
            Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(2)
            lValorComis = lValorComis + lObjstcCorr.Per_Corr
            If pobjstcPed.ColecaoPedCorr.Count > 2 Then
                Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(3)
                lValorComis = lValorComis + lObjstcCorr.Per_Corr
            End If
        End If
    End If

    If lValorComis = 0 Then
        Print #pNumArqDados, "Prêmio Net: R$"
    Else
        Print #pNumArqDados, "Prêmio Líquido: R$"
    End If
    GeraCYX_LMI lLinCXY, "Premio", lPosY, 0.5, False, "S", lNumArqCXY

    lAuxiliar2 = Format$(pobjstcPed.VAL_PR, "standard")
    lAuxiliar2 = gfPreencheBranco(lAuxiliar2, 20)
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    GeraCYX_LMI lLinCXY, "Premio", lPosY, 4.5, True, "N", lNumArqCXY    '

    GeraCYX_LMI lLinCXY, "Linha Vazia", lPosY, 0.5, True, "N", lNumArqCXY    '

    'demnostrativo de prêmio

    'Controle de página
    lNumPagina = 1
    lTotPagina = 1

    If lValorComis <> 0 Then
        lAuxiliar2 = String(70, "-") & "DEMONSTRATIVO DO PRÊMIO" & String(70, "-")
        Print #pNumArqDados, lAuxiliar2
        GeraCYX_LMI lLinCXY, "Título de demonstrativo de prêmio", lPosY, 0.5, True, "S", lNumArqCXY

        lAuxiliar2 = "Prêmio Líquido" & Space(20) & "Custo da Apólice" & Space(20) & "IOF (7,38)" & Space(20) & "Prêmio Total Vista"
        Print #pNumArqDados, lAuxiliar2
        GeraCYX_LMI lLinCXY, "Título de demonstrativo de prêmio 2", lPosY, 0.5, True, "S", lNumArqCXY

        '013. Totalizar - Prêmio Líquido
        lAuxiliar2 = Format$(pobjstcPed.VAL_PR, "standard")
        lTamanho = Len(lAuxiliar2)
        If lTamanho < 14 Then
            lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
        End If
        Print #pNumArqDados, lAuxiliar2
        GeraCYX_LMI lLinCXY, "Prêmio líquido", lPosY, 0.5, False, "N", lNumArqCXY


        '014. Custo da Apólice
        lAuxiliar2 = Format$(pobjstcPed.Val_Cust_Emis, "standard")
        lTamanho = Len(lAuxiliar2)
        If lTamanho < 14 Then
            lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
        End If
        Print #pNumArqDados, lAuxiliar2
        GeraCYX_LMI lLinCXY, "Custo de apólice", lPosY, 5, False, "N", lNumArqCXY

        '015. Valor IOF.
        lAuxiliar2 = Format$(pobjstcPed.Val_IOF, "standard")
        lTamanho = Len(lAuxiliar2)
        If lTamanho < 14 Then
            lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
        End If
        Print #pNumArqDados, lAuxiliar2
        GeraCYX_LMI lLinCXY, "IOF", lPosY, 9, False, "N", lNumArqCXY

        '016. Prêmio total a vista.
        lAuxiliar2 = Format$(pobjstcPed.gValorTotal, "standard")
        lTamanho = Len(lAuxiliar2)
        If lTamanho < 14 Then
            lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
        End If
        Print #pNumArqDados, lAuxiliar2
        GeraCYX_LMI lLinCXY, "Prêmio total", lPosY, 14, True, "N", lNumArqCXY


        lAuxiliar2 = String(62, "-") & " DEMONSTRATIVO DE PARCELAMENTO " & String(62, "-")
        Print #pNumArqDados, lAuxiliar2
        GeraCYX_LMI lLinCXY, "Título de demonstrativo de parcelamento", lPosY, 0.6, True, "S", lNumArqCXY

'        lAuxiliar2 = "                   Primeira parcela a vista                                                                      Primeira parcela a 30 dias"
'        Print #pNumArqDados, lAuxiliar2
'        GeraCYX_LMI lLinCXY, "Título de demonstrativo de parcelamento 2", lPosY, 0.6, True, "S", lNumArqCXY
'
'        lAuxiliar2 = String(80, "-") & Space(22) & String(82, "-")
'        Print #pNumArqDados, lAuxiliar2
'        GeraCYX_LMI lLinCXY, "Título de demonstrativo de parcelamento 3", lPosY, 0.6, True, "S", lNumArqCXY
'
'        lAuxiliar2 = "Parc              1ª parcela                 Demais           Juros                        1ª parcela            Demais           Juros "
'        Print #pNumArqDados, lAuxiliar2
'        GeraCYX_LMI lLinCXY, "Título de demonstrativo de parcelamento 4", lPosY, 0.6, True, "S", lNumArqCXY


        'Calcula o fracionamento
        Set lobjstcPedLoc = pobjstcPed.ColecaoPedLoc.Item(1)
        With pobjstcPed
            Call gM46V998.gfCalFracionamento(pobjstcPed, lobjstcPedLoc.Tip_Emissao, _
                                             lValorParc, lMensagem, gM46V111.gProJurIOF, .Cod_Reg, .Val_IOF, .Val_Parc_Min)
        End With

        '018. Formas de pagamento - primeira parcela a 30 dias.
        'Ficha 900665
        lOrdem = 1
        Dim i As Integer
        For i = 1 To gM46V999.gQtdMaxComp
            If i = 1 Then
                lAuxiliar2 = "                   Primeira parcela boleto a vista                                                                      Primeira parcela boleto a 30 dias"
            Else
                lAuxiliar2 = "                   Primeira parcela débito a vista                                                                      Primeira parcela débito a 30 dias"
            End If
            
            Print #pNumArqDados, lAuxiliar2
            
            GeraCYX_LMI lLinCXY, "Título de demonstrativo de parcelamento 2", lPosY, 0.6, True, "S", lNumArqCXY, pNumArqDados, pobjstcPed
    
            lAuxiliar2 = String(80, "-") & Space(22) & String(82, "-")
            Print #pNumArqDados, lAuxiliar2
            GeraCYX_LMI lLinCXY, "Título de demonstrativo de parcelamento 3", lPosY, 0.6, True, "S", lNumArqCXY, pNumArqDados, pobjstcPed
    
            lAuxiliar2 = "Parc              1ª parcela                 Demais           Juros                        1ª parcela            Demais           Juros "
            Print #pNumArqDados, lAuxiliar2
            GeraCYX_LMI lLinCXY, "Título de demonstrativo de parcelamento 4", lPosY, 0.6, True, "S", lNumArqCXY, pNumArqDados, pobjstcPed

        For lI = 1 To 12
            If lI > pobjstcPed.Qtd_Parc Then
                Exit For
            Else
                'a) Ordem da parcela.
                lOrdem = Format$(lI, "00")
                'b) Valor da primeira parcela.
                lAuxiliar1 = lValorParc(lI, 1, i)
                If lAuxiliar1 <> 0 Then
                    lPriParcela = Format$(lAuxiliar1, "standard")
                    lTamanho = Len(lPriParcela)
                    If lTamanho < 10 Then
                        lPriParcela = Space(10 - lTamanho) & lPriParcela
                    End If
                    'c) Valor das demais parcelas.
                    lAuxiliar1 = lValorParc(lI, 2, i)
                    lDemParcelas = Format$(lAuxiliar1, "standard")
                    lTamanho = Len(lDemParcelas)
                    If lTamanho < 9 Then
                        lDemParcelas = Space(9 - lTamanho) & lDemParcelas
                    End If
                    'd) Juros.
                    Set lobjstcJur = New stcA46V711B
                    On Error Resume Next
                    Set lobjstcJur = pobjstcPed.ColecaoPedJuros.Item(pobjstcPed.NOSSO_NUMERO & Format(lI, "00") & Format(i, "00"))
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
                    lAuxiliar2 = Format(lOrdem, "00") & "   " & lPriParcela & "   " & lDemParcelas & "  " & lJuros
                Else
                    lAuxiliar2 = "- "
                End If
                If lValorParc(lI, 3, i) <> 0 Then
                    If lAuxiliar2 = "- " Then   'Ficha 1000099 - Planos 2131 a 2138 Corretor Schroeder - para imprimir corretamente parc a 30dias
                        lAuxiliar2 = Format(lOrdem, "00") & Space(31)
                    End If
                    'b) Valor da primeira parcela.
                    lAuxiliar1 = lValorParc(lI, 3, i)
                    If lAuxiliar1 <> 0 Then
                        lPriParcela = Format$(lAuxiliar1, "standard")
                        lTamanho = Len(lPriParcela)
                        If lTamanho < 9 Then
                            lPriParcela = Space(9 - lTamanho) & lPriParcela
                        End If
                        'c) Valor das demais parcelas.
                        lAuxiliar1 = lValorParc(lI, 4, i)
                        lDemParcelas = Format$(lAuxiliar1, "standard")
                        lTamanho = Len(lDemParcelas)
                        If lTamanho < 9 Then
                            lDemParcelas = Space(9 - lTamanho) & lDemParcelas
                        End If
                        'd) Juros.
                        Set lobjstcJur = New stcA46V711B
                        On Error Resume Next
                        Set lobjstcJur = pobjstcPed.ColecaoPedJuros.Item(pobjstcPed.NOSSO_NUMERO & Format(lI, "00") & Format(i, "00"))
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
                        lAuxiliar2 = lAuxiliar2 & "     " & lPriParcela & "  " & lDemParcelas & "  " & lJuros
                    Else
                        lAuxiliar2 = lAuxiliar2 & " "
                    End If
                Else
                    lAuxiliar2 = lAuxiliar2 & " "
                End If
                    
                    If lAuxiliar2 <> "-  " Then
                        Print #pNumArqDados, lAuxiliar2
                        GeraCYX_LMI lLinCXY, "Parcela ", lPosY, 0.6, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed, "COURIER NEW"
                    End If
                End If
            Next lI
        Next i

    End If

    If lQtdCobert >= 13 Then
        lTotPagina = lTotPagina + Int(pobjstcPed.ColecaoPedLoc.Count / 2)
        If pobjstcPed.ColecaoPedLoc.Count Mod 2 > 0 Then
            lTotPagina = lTotPagina + 1
        End If
    ElseIf lQtdCobert >= 6 Then
        lTotPagina = lTotPagina + Int(pobjstcPed.ColecaoPedLoc.Count / 4)
        If pobjstcPed.ColecaoPedLoc.Count Mod 4 > 0 Then
            lTotPagina = lTotPagina + 1
        End If
    Else
        lTotPagina = lTotPagina + Int(pobjstcPed.ColecaoPedLoc.Count / 6)
        If pobjstcPed.ColecaoPedLoc.Count Mod 6 > 0 Then
            lTotPagina = lTotPagina + 1
        End If
    End If

    GeraCYX_LMI lLinCXY, "Linha Vazia", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed

    '013. Label Riscos Cobertos.
    Print #pNumArqDados, "Riscos Cobertos:"
    GeraCYX_LMI lLinCXY, "Label Riscos Cobertos", lPosY, 0.5, False, "S", lNumArqCXY, pNumArqDados, pobjstcPed

    '014. Descrição Riscos Cobertos.
    lAuxiliar2 = " "
    Select Case pobjstcPed.Cod_Ramo
    Case e_CodRamo_Condominio
        lAuxiliar2 = "Conforme Condições Gerais do Produto Yasuda Condomínio"
    Case e_CodRamo_Empresarial
        lAuxiliar2 = "Conforme Condições Gerais do Produto Yasuda Empresarial"
    Case e_CodRamo_Residencial
        lAuxiliar2 = "Conforme Condições Gerais do Produto Yasuda Residencial"
    End Select

    Print #pNumArqDados, lAuxiliar2
    GeraCYX_LMI lLinCXY, "Label Riscos Cobertos", lPosY, 4.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed

    GeraCYX_LMI lLinCXY, "Linha Vazia", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed

    '015. Resumo: quantidade de itens.
    Print #pNumArqDados, "Locais de Risco:"
    GeraCYX_LMI lLinCXY, "Locais de Risco", lPosY, 0.5, False, "S", lNumArqCXY, pNumArqDados, pobjstcPed

    Print #pNumArqDados, "(" & pobjstcPed.ColecaoPedLoc.Count & " locais) conforme relação anexa"
    GeraCYX_LMI lLinCXY, "Locais de Risco", lPosY, 4.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed

    GeraCYX_LMI lLinCXY, "Linha Vazia", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed

    '016. Valor em Risco.
    Print #pNumArqDados, "Valor em Risco:"
    GeraCYX_LMI lLinCXY, "Valor em Risco", lPosY, 0.5, False, "S", lNumArqCXY, pNumArqDados, pobjstcPed

    Print #pNumArqDados, " Conforme relação anexa"
    GeraCYX_LMI lLinCXY, "Valor em Risco", lPosY, 4.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed

    GeraCYX_LMI lLinCXY, "Linha Vazia", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed

    '017 Ocupação.
    Print #pNumArqDados, "Ocupação:"
    GeraCYX_LMI lLinCXY, "Ocupacao", lPosY, 0.5, False, "S", lNumArqCXY, pNumArqDados, pobjstcPed

    Print #pNumArqDados, " Conforme relação anexa"
    GeraCYX_LMI lLinCXY, "Ocupacao", lPosY, 4.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed


    GeraCYX_LMI lLinCXY, "Linha Vazia", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed

    '018 Franquia.
    Print #pNumArqDados, "Franquias/Período indenitário:"
    GeraCYX_LMI lLinCXY, "Franquia", lPosY, 0.5, True, "S", lNumArqCXY, pNumArqDados, pobjstcPed

    'Ficha 1000655 (item 8 inclusão de identificação do usuário no rodapé da página)
    Print #pNumArqDados, pobjstcPed.Cod_User_Liber & " - " & gM46V999.gfForData(2, pobjstcPed.Dat_Liber) & " - " & Mid(Format(pobjstcPed.Hor_Liber, "0000"), 1, 2) & ":" & Mid(Format(pobjstcPed.Hor_Liber, "0000"), 3, 2)
    GeraCYX_LMI lLinCXY, "Detalhes usuário", 28.3, 0.6, False, "N", lNumArqCXY, pNumArqDados, pobjstcPed, , True

    For lI = 1 To lQtdCobert
        If vCobert(3, lI) > 0 And (vCobert(4, lI) = 0 Or gM46V999.gCodProduto < 1080) Then 'Tem franquia
            lAuxiliar2 = vCobert(1, lI) & ": "  'Descrição da cobertura
            Print #pNumArqDados, lAuxiliar2
            GeraCYX_LMI lLinCXY, "Franquia", lPosY, 0.5, False, "N", lNumArqCXY, pNumArqDados, pobjstcPed

            lAuxiliar2 = ""
            If vCobert(5, lI) > 0 Then
                lAuxiliar2 = vCobert(5, lI) & "% dos prej.indeniz., min: "
            End If
            lAuxiliar2 = lAuxiliar2 & "R$ " & Format$(vCobert(3, lI), "standard")
            Print #pNumArqDados, lAuxiliar2
            GeraCYX_LMI lLinCXY, "Franquia", lPosY, 13, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
        ElseIf vCobert(4, lI) > 0 Then  'Pi indenitário
            lAuxiliar2 = vCobert(1, lI) & ": "    'Descrição da cobertura
            Print #pNumArqDados, lAuxiliar2
            GeraCYX_LMI lLinCXY, "Franquia", lPosY, 0.5, False, "N", lNumArqCXY, pNumArqDados, pobjstcPed

            If vCobert(3, lI) > 0 Then
                lAuxiliar2 = vCobert(3, lI) & " dias"
            Else
                lAuxiliar2 = "PI = " & vCobert(4, lI) & " meses"
            End If
            
            Call mpPrazoTabCoefPI(pobjstcPed, CInt(vCobert(0, lI)), CInt(vCobert(4, lI)), lPrazo)
            If lPrazo <> 0 Then
                lAuxiliar2 = lAuxiliar2 & ": " & lPrazo & " dias"
            End If
            Print #pNumArqDados, lAuxiliar2
            GeraCYX_LMI lLinCXY, "Franquia", lPosY, 13, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
        ElseIf vCobert(6, lI) = "*" Then  ' Asterístico (™ by Luana)
            lAuxiliar2 = vCobert(1, lI) & ": "    'Descrição da cobertura
            Print #pNumArqDados, lAuxiliar2
            GeraCYX_LMI lLinCXY, "Franquia", lPosY, 0.5, False, "N", lNumArqCXY, pNumArqDados, pobjstcPed

            lAuxiliar2 = "*"
            Print #pNumArqDados, lAuxiliar2
            GeraCYX_LMI lLinCXY, "Franquia", lPosY, 13, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
        End If
    Next

    For lI = 0 To 50
        If lDescFranquias(lI) = "" Then Exit For
        lAuxiliar2 = lDescFranquias(lI)
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        GeraCYX_LMI lLinCXY, "Franquia", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
    Next lI

    GeraCYX_LMI lLinCXY, "Linha Vazia", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed

    '019 Cláusulas e Condições.
    Print #pNumArqDados, "Cláusulas e Condições:"
    GeraCYX_LMI lLinCXY, "Clausulas e Condicoes:", lPosY, 0.5, False, "S", lNumArqCXY, pNumArqDados, pobjstcPed

    Print #pNumArqDados, "Ratificam-se as seguintes Cláusulas e Condições:"
    GeraCYX_LMI lLinCXY, "Clausulas e Condicoes", lPosY, 4.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed

    lAuxiliar2 = " "
    Select Case pobjstcPed.Cod_Ramo
    Case e_CodRamo_Condominio
        If pobjstcPed.Cod_Produto >= 1080 Then
            lAuxiliar2 = "a) Condições Gerais do Produto Condomínio (versão 10/2014)"
        Else
            lAuxiliar2 = "a) Condições Gerais do Produto Yasuda Condomínio (versão 09/2008)"
        End If
            
        Print #pNumArqDados, lAuxiliar2
        GeraCYX_LMI lLinCXY, "Clausulas e Condicoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
        lI = 98
    Case e_CodRamo_Empresarial
        If pobjstcPed.Cod_Produto >= 1080 Then
            lAuxiliar2 = "a) Condições Gerais do Produto Empresarial (versão 10/2014)"
        Else
            lAuxiliar2 = "a) Condições Gerais do Produto Yasuda Empresarial (versão 06/2009)"
        End If
            
        Print #pNumArqDados, lAuxiliar2
        GeraCYX_LMI lLinCXY, "Clausulas e Condicoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
        '1200029 - Clausula 315 e Texto LMI
        '''            lAuxiliar2 = "b) Cláusula 315- LMI Único "
        '''            Print #pNumArqDados, lAuxiliar2
        '''            GeraCYX_LMI lLinCXY, "Clausulas e Condicoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
        lI = 98
    Case e_CodRamo_Residencial
        If pobjstcPed.Cod_Produto >= 1080 Then
            lAuxiliar2 = "a) Condições Gerais do Produto Residencial (versão 10/2014)"
        Else
            lAuxiliar2 = "a) Condições Gerais do Produto Yasuda Residencial (versão 09/2008)"
        End If
            
        Print #pNumArqDados, lAuxiliar2
        GeraCYX_LMI lLinCXY, "Clausulas e Condicoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
        lI = 98
    End Select

    If pobjstcPed.Cod_Ramo = e_CodRamo_Residencial Or pobjstcPed.Cod_Ramo = e_CodRamo_Condominio Then
        lAuxiliar2 = Chr(lI) & ") CLÁUSULA PARTICULAR L.M.I. Único - Declara-se para os devidos fins e efeitos que, este seguro foi contratado sob "
        lI = lI + 1
        Print #pNumArqDados, lAuxiliar2
        GeraCYX_LMI lLinCXY, "Clausulas e Condicoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
        lAuxiliar2 = "a forma de LMI Único para todos os locais segurados especificados na presente apólice. Ao contrário do que possa constar "
        Print #pNumArqDados, lAuxiliar2
        GeraCYX_LMI lLinCXY, "Clausulas e Condicoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
        lAuxiliar2 = "nas condições gerais do presente contrato de seguro, fica entendido e acordado que em caso de eventual sinistro, a "
        Print #pNumArqDados, lAuxiliar2
        GeraCYX_LMI lLinCXY, "Clausulas e Condicoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
        lAuxiliar2 = "indenização não poderá exceder ao LMI de cada cobertura contratada, limitado ao VR declarado para cada local."
        Print #pNumArqDados, lAuxiliar2
        GeraCYX_LMI lLinCXY, "Clausulas e Condicoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
    End If

    lListClausula = ""
    Dim ArrClausulasJaUsadas As String
    ArrClausulasJaUsadas = ""

    For Each lobjstcPedLoc In pobjstcPed.ColecaoPedLoc
        If lobjstcPedLoc.ColecaoPedClau.Count > 0 Then
            ' aplicar as regras de exclusão
            Set lObjstcPedLocTmp = lobjstcPedLoc
            Dim asj As Integer
            asj = 1
            For asj = 1 To lObjstcPedLocTmp.ColecaoPedClau.Count
                If lObjstcPedLocTmp.ColecaoPedClau.Count = 0 Then Exit For

                Set lobjstcClau = lObjstcPedLocTmp.ColecaoPedClau(asj)
                If lobjstcClau.Cod_Clau >= 310 Or (lobjstcClau.Cod_Ramo = e_CodRamo_Residencial And lobjstcClau.Cod_Clau > 300) Or pobjstcPed.LMI_Unico = 1 Then
                    If Not lListClausula Like "*;" & lobjstcClau.Cod_Clau & ";*" And InStr(ArrClausulasJaUsadas, lobjstcClau.Cod_Clau) = 0 Then
                        ' Nada
                        ArrClausulasJaUsadas = ArrClausulasJaUsadas & lobjstcClau.Cod_Clau & ","
                    Else    'Excluir
                        lObjstcPedLocTmp.ColecaoPedClau.Remove (asj)
                        asj = asj - 1
                    End If
                Else    ' Excluir
                    lObjstcPedLocTmp.ColecaoPedClau.Remove (asj)
                    asj = asj - 1
                End If
                If lobjstcClau.Cod_Clau = 400 Or lobjstcClau.Cod_Clau = 401 Or lobjstcClau.Cod_Clau = 402 Then
                    lTemAssistencia = True
                End If
                If asj = lObjstcPedLocTmp.ColecaoPedClau.Count Then Exit For


            Next

            Titulo = "Cláusulas"

            If lObjstcPedLocTmp.ColecaoPedClau.Count > 0 Then
                If gM46V999.gCodProduto >= 1014 Then
                    Call mpImpClausulas(pNumArqDados, lLinCXY, lPosY, 0.5, lNumArqCXY, pobjstcPed, lObjstcPedLocTmp, Titulo, True)
                End If
            End If

            '        For Each lobjstcClau In lObjstcPedLoc.ColecaoPedClau
            '            If lobjstcClau.Cod_Clau >= 310 Or (lobjstcClau.Cod_Ramo = e_CodRamo_Residencial And lobjstcClau.Cod_Clau > 300) Or pobjstcPed.LMI_Unico = 1 Then
            '                If Not lListClausula Like "*;" & lobjstcClau.Cod_Clau & ";*" Then
            '                    lListClausula = lListClausula & ";" & lobjstcClau.Cod_Clau & ";"
            '                    lAuxiliar2 = Chr(lI) & ") Cláusula " & lobjstcClau.Cod_Clau & " - " & lobjstcClau.gDescClausula
            '                    Print #pNumArqDados, lAuxiliar2
            '                    GeraCYX_LMI lLinCXY, "Clausulas e Condicoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
            '                    lI = lI + 1
            '                End If
            '            End If
            '        Next lobjstcClau
        End If
    Next

    GeraCYX_LMI lLinCXY, "Linha Vazia", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed

    '021 Observações .
    lI = 0
    lTam = 95
    lTemObs = False
    TemInspecaoRisco = False
    lAuxiliar2 = Empty
    
    If gM46V111.gNum_Pi = pobjstcPed.NUM_PI Then    'Se estiver imprimindo o mesmo que está sendo calculado
        For Each lobjstcPedLoc In pobjstcPed.ColecaoPedLoc

            'Set lobjstcPedLoc = pobjstcPed.ColecaoPedLoc.Item(1)
            lAuxiliar2 = lobjstcPedLoc.objstcPedObs.Observacao
            If gM46V999.gfPreenchido(lAuxiliar2) Then
                lAuxiliar2 = lAuxiliar2 & "&%"
            End If

            lObs = Split(lAuxiliar2, "&%")
            For lCount = 0 To UBound(lObs) - 1
                lAuxiliar2 = lObs(lCount)
                If lAuxiliar2 Like "*OBRIGATORIO A REALIZACAO DE INSPECAO DO RISCO*" Then
                    lMsgInspecao = lAuxiliar2
                    lAuxiliar2 = ""
                End If
                If lAuxiliar2 Like "*ENVIO DE COPIA DA APOLICE DA CONGENERE*" Then
                    lTextoCong = lAuxiliar2
                    lAuxiliar2 = ""
                End If
                If InStr(UCase(lAuxiliar2), "OBRIGATORIA A REALIZACAO PREVIA DE INSPECAO DO RISCO") Then
                    TemInspecaoRisco = True
                    lAuxiliar2 = ""
                End If
                
                If InStr(UCase(lAuxiliar2), "INSPECAO") Then
                    TemInspecaoRisco = True
                End If

                If gM46V999.gfPreenchido(lAuxiliar2) Then
                    If Len(lAuxiliar2) < lTam Then
                        ReDim Preserve lDescImportante(lI)
                        lDescImportante(lI) = lAuxiliar2
                        lTemObs = True
                        lI = lI + 1
                    Else
                        lPos = 1
                        Do
                            ReDim Preserve lDescImportante(lI)
                            lTemObs = True
                            Aux = 0
                            If Len(Mid(lAuxiliar2, lPos)) > lTam Then
                                Aux = InStrRev(Mid(lAuxiliar2, lPos, lTam), " ")
                                'Sandra
                                If Aux > 0 And Aux < lTam Then
                                    Aux = lTam - Aux
                                Else
                                    Aux = 0
                                End If
                            End If
                            lDescImportante(lI) = Mid(lAuxiliar2, lPos, lTam - Aux)
                            lI = lI + 1
                            'Sandra
                            If (Len(Mid(lAuxiliar2, lPos, lTam)) <= lTam) And Aux = 0 Then Exit Do
                            lPos = lPos + (lTam - Aux)
                        Loop
                    End If
                End If
            Next lCount
        Next lobjstcPedLoc

        If TemInspecaoRisco Then
            ReDim Preserve lDescImportante(lI)
            lDescImportante(lI) = "OBRIGATORIA A REALIZACAO PREVIA DE INSPECAO DO RISCO PARA OS LOCAIS RELACIONADOS"
            lI = lI + 1
            ReDim Preserve lDescImportante(lI)
            lDescImportante(lI) = "COM A EXPRESSAO ""INSP""."
            lI = lI + 1
            lTemObs = True
        End If
    End If
    
    Print #pNumArqDados, "Observações:"
    GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "S", lNumArqCXY, pNumArqDados, pobjstcPed
    lContador = 1
    lAuxiliar2 = lContador & "-Prêmios, taxas, franquias e demais condições sujeitas à alteração mediante inspeção no local segurado, e/ou caso sejam"
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed

    lAuxiliar2 = "constatados fatores e/ou informações divergentes daquelas fornecidas e/ou atividades agravantes provisórias e/ou definitivas."
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed

    'Ficha 1000742 - Não exibir especificação Item 2 para Corretor Volkswagem
    blnCorretorVolks = False
    For Each lObjstcCorr In pobjstcPed.ColecaoPedCorr
        If InStr(1, "0008001, 0008002, 0008003, 0008004, 0008005, 0008006, 0008008, 0008011, 0008014, 0008028", lObjstcCorr.Cod_Corr, vbTextCompare) > 0 Then
            blnCorretorVolks = True
            If lConcessionaria = False Then blnCorretorVolks = False
        End If
    Next
    lContador = lContador + 1
    'Ficha 1000742 - Não exibir especificação Item 2 para Corretor Volkswagem
    If blnCorretorVolks = False Then
        lAuxiliar2 = lContador & "-Acham-se excluídas da presente cotação as construções mistas, inferiores, armazéns infláveis, galpão de vinilona e"
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed

        lAuxiliar2 = "assemelhados e seus respectivos conteúdos, porventura existentes no local segurado."
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
    End If

    lContador = lContador + 1
    If lValorComis = 0 Then
        lAuxiliar2 = lContador & "-Prêmio acima desprovido de Comissão de Corretagem"
    Else
        lAuxiliar2 = lContador & "-Prêmio acima com carregamento de Comissão de Corretagem"
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed

    lContador = lContador + 1
    lAuxiliar2 = lContador & "-A presente cotação é apresentada apenas como parâmetro de custos, e não implica no reconhecimento da aceitação do risco,"
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed

    lAuxiliar2 = "a qual dependerá da análise do mesmo,após a respectiva inspeção,bem como as taxas,franquias e condições."
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
    lContador = lContador + 1
    lAuxiliar2 = lContador & "-O registro deste plano na SUSEP não implica, por parte da Autarquia, incentivo ou recomendação a sua comercialização."
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
    lContador = lContador + 1
    lAuxiliar2 = lContador & "-O segurado poderá consultar a situação cadastral de seu corretor de seguros, no site www.susep.gov.br, por meio de seu"
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed

    lAuxiliar2 = "registro na SUSEP, nome completo, CNPJ ou CPF."
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
    lContador = lContador + 1
    lAuxiliar2 = lContador & "-Em havendo outros seguros efetuados do mesmo local e mesmo segurado à Yasuda Seguros deverá ser comunicado,"
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed

    lAuxiliar2 = "tendo em vista acumulo de risco."
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
    lContador = lContador + 1
    lAuxiliar2 = lContador & "-Cotação válida por 15 dias."
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed



    'thais
    If lTemAssistencia = False Then
        lContador = lContador + 1
        Select Case pobjstcPed.Cod_Ramo
        Case e_CodRamo_Condominio
            lAuxiliar2 = lContador & "-Assistência 24 Horas Gratuita."
        Case e_CodRamo_Empresarial
            lAuxiliar2 = lContador & "-Assistência 24 Horas Gratuita."
        Case e_CodRamo_Residencial
            lAuxiliar2 = lContador & "-Assistência Domiciliar 24 Horas Gratuita (exceto para residências eventuais)."
        End Select
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    End If



    '1101065 - Mensagem LMI
    GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
    lContador = lContador + 1
    lAuxiliar2 = lContador & "-Esta cotação foi elaborada sob a forma de LMI Único. Verifique os sub-limites estabelecidos na relação dos locais de risco"
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
    lAuxiliar2 = "para cada endereço segurado e, caso não localize ou ainda, haja alguma divergência quanto aos valores e/ou locais segurados,"
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
    lAuxiliar2 = "entre em contato com seu corretor de seguros e solicite uma nova cotação."
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    Dim lobjstcTextoEspec
    If pobjstcPed.ColecaoTextoEspec.Count > 0 Then
        For Each lobjstcTextoEspec In pobjstcPed.ColecaoTextoEspec
            'If lobjstcTextoEspec.Cod_Espec <> 8500 Then
            If gM46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin1) Then
                GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
                lAuxiliar2 = lobjstcTextoEspec.Espec_lin1
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            End If
            If gM46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin2) Then
                GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
                lAuxiliar2 = lobjstcTextoEspec.Espec_lin2
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            End If
            If gM46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin3) Then
                GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
                lAuxiliar2 = lobjstcTextoEspec.Espec_lin3
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            End If
            If gM46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin4) Then
                GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
                lAuxiliar2 = lobjstcTextoEspec.Espec_lin4
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            End If
            If gM46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin5) Then
                GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
                lAuxiliar2 = lobjstcTextoEspec.Espec_lin5
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            End If
            If gM46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin6) Then
                GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
                lAuxiliar2 = lobjstcTextoEspec.Espec_lin6
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            End If
            If gM46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin7) Then
                GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
                lAuxiliar2 = lobjstcTextoEspec.Espec_lin7
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            End If
            If gM46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin8) Then
                GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
                lAuxiliar2 = lobjstcTextoEspec.Espec_lin8
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            End If
            If gM46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin9) Then
                GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
                lAuxiliar2 = lobjstcTextoEspec.Espec_lin9
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            End If
            If gM46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin10) Then
                GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
                lAuxiliar2 = lobjstcTextoEspec.Espec_lin10
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            End If
            If gM46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin11) Then
                GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
                lAuxiliar2 = lobjstcTextoEspec.Espec_lin11
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            End If
            If gM46V999.gfPreenchido(lobjstcTextoEspec.Espec_lin12) Then
                GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
                lAuxiliar2 = lobjstcTextoEspec.Espec_lin12
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            End If
            'End If
        Next lobjstcTextoEspec
    End If


    If pobjstcPed.Cod_Ramo = e_CodRamo_Empresarial And lConcessionaria = True Then
        Select Case pobjstcPed.objstcCtrlEmis.Cod_Corr
        Case 8001, 8002, 8003, 8004, 8005, 8006, 8008, 8011, 8014, 8028
            GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
            lAuxiliar2 = "Mercadorias e matérias-primas de terceiros em poder do segurado: seus valores deverão ser considerados na composição "
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
            lAuxiliar2 = "do Valor Risco Total Declarado, devendo ser observado os bens excluídos das Condições Gerais do Produto Multi"
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
            lAuxiliar2 = "Seguro Empresarial;"
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        End Select
    End If

    'Ficha 1000307
    If gObrRegistro = True Then
        If pobjstcPed.Cod_Ramo = e_CodRamo_Empresarial And pobjstcPed.Tip_Pessoa = e_TipPessoa_Fisica And gM46V999.gTipCobertura <> e_TipCobert_Predio Then
            GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
            lAuxiliar2 = "Seguro de Pessoa Física: Na proposta será obrigatório informar o número do registro profissional no respectivo Conselho "
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
            lAuxiliar2 = "da Categoria. Ex.: OAB, CRM, CRO etc. "
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        End If
    End If

    If lTemObs = True Then
        GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
    Else
        GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed, , True
    End If

    If gM46V111.gNum_Pi = pobjstcPed.NUM_PI Then    'Se estiver imprimindo o mesmo que está sendo calculado
        If lTemObs = True Then
            For lI = 0 To UBound(lDescImportante)
                lAuxiliar2 = lDescImportante(lI)
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
                GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
            Next lI
        End If
    End If

    If gM46V999.gfPreenchido(lMsgInspecao) Then
        lAuxiliar2 = lMsgInspecao
        If Len(lAuxiliar2) < lTam Then
            Print #pNumArqDados, lMsgInspecao
            GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
        Else
            lPos = 1
            Do
                Aux = 0
                If Len(Mid(lAuxiliar2, lPos)) > lTam Then
                    Aux = InStrRev(Mid(lAuxiliar2, lPos, lTam), " ")
                    If Aux < lTam Then
                        Aux = lTam - Aux
                    Else
                        Aux = 0
                    End If
                End If
                Print #pNumArqDados, Mid(lAuxiliar2, lPos, lTam - Aux)
                GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
                If (Len(Mid(lAuxiliar2, lPos, lTam)) < lTam) And Aux = 0 Then Exit Do
                lPos = lPos + (lTam - Aux)
            Loop
        End If
    End If
    If gM46V999.gfPreenchido(lTextoCong) Then
        lAuxiliar2 = lTextoCong
        If Len(lAuxiliar2) < lTam Then
            Print #pNumArqDados, lTextoCong
            GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
        Else
            lPos = 1
            Do
                Aux = 0
                If Len(Mid(lAuxiliar2, lPos)) > lTam Then
                    Aux = InStrRev(Mid(lAuxiliar2, lPos, lTam), " ")
                    If Aux < lTam Then
                        Aux = lTam - Aux
                    Else
                        Aux = 0
                    End If
                End If
                Print #pNumArqDados, Mid(lAuxiliar2, lPos, lTam - Aux)
                GeraCYX_LMI lLinCXY, "Observacoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
                If (Len(Mid(lAuxiliar2, lPos, lTam)) < lTam) And Aux = 0 Then Exit Do
                lPos = lPos + (lTam - Aux)
            Loop
        End If
    End If

    If lTemDesc = True Then
        Print #pNumArqDados, "PRÊMIO CALCULADO COM DESCONTO COMERCIAL"
        GeraCYX_LMI lLinCXY, "Clausulas e Condicoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
    End If

    'Texto Observação de Franquia
    Dim Linha As Variant
    
    For Each lobjstcPedLoc In pobjstcPed.ColecaoPedLoc
        For Each lobjstcCob In lobjstcPedLoc.ColecaoPedCobert
            If lobjstcCob.ColecaoNovaFranquia.Count > 0 Then
                If lobjstcCob.ColecaoNovaFranquia.Item(1).Texto_Observacoes <> Empty Then
                    Print #pNumArqDados, " "
                    GeraCYX_LMI lLinCXY, "Clausulas e Condicoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
                    
                    Linha = mfSplitEX(lobjstcCob.ColecaoNovaFranquia.Item(1).Texto_Observacoes, Asc(vbCrLf), 110)
                    
                    For i = 0 To UBound(Linha)
                        Print #pNumArqDados, Linha(i)
                        GeraCYX_LMI lLinCXY, "Clausulas e Condicoes", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
                    Next i
                End If
            End If
        Next
    Next

    'thais - lmi
    For Each lobjstcPedDizer In pobjstcPed.ColecaoTextoEndo
        Print #pNumArqDados, lobjstcPedDizer.Dsc_Dizer_Endo
        GeraCYX_LMI lLinCXY, "Dizeres de endosso", lPosY, 0.5, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
    Next lobjstcPedDizer

    'Rodape

    Print #pNumArqDados, "Página: " & lNumPagina & "/@@TOTALPAG@@"
    GeraCYX_LMI lLinCXY, "Pagina", 28.5, 18, True, "N", lNumArqCXY, pNumArqDados, pobjstcPed
    lNumPagina = lNumPagina + 1
Fim:
    Print #lNumArqCXY, " "
    Close #lNumArqCXY

    mfCriDados_COT_LMI = True

    Exit Function
ControleErro:
    Call gM46V999.gpGraLog(1, "mfCriDados_COT_LMI ERRO - " & Err & " - " & Error)

End Function



Public Function gfImpDocumento(ByVal pOpcao As Byte, ByVal pFormulario As String, _
                               ByVal pImagem As String, ByVal pobjstcPed As stcA46V702B, _
                               pobjstcPedLoc As stcA46V708B, pPagDe As Integer, _
                               pPagAte As Integer, pNomeArqIST As String) As Boolean
'Função: imprime documento.

'Parâmetros de entrada...pOpcao...........0 = visualizar impressão.
'                                         1 = imprimir.
'                                         2 = criar arquivo imagem.
'                        pFormulario......Nome do formulário.
'                        pImagem..........Nome para o arquivo imagem a ser criado.
'                        lObjP0042200.....Registros de item.
    Dim lVerifOverf As Boolean
    Dim lI As Integer
    Dim lArqDados As String         'Endereço do arquivo de dados.
    Dim lEndConfig As String          'Endereço do arquivo de configuração.
    Dim lEndImagem As String          'Endereço do arquivo de imagem.
    Dim lNumArquivo As Integer          'Número do arquivo.
    Dim lRefConfig As String          'Número de referência do arquivo de configuração.
    Dim lRefImagem As String          'Número de referência do arquivo de imagem.
    Dim lResultado As Boolean         'Resultado das funções de criação de dados.
    Dim lPossuiObs As Boolean         'Possui observação?
    Dim lobjstcLoc As stcA46V708B     'Objeto de itens.
    Dim lArrayArq
    Dim lCount As Integer
    Dim lPos As Integer
    Dim lArqIST As Boolean
    Dim lRegistro As String
    Dim lNumArqIST As Integer

Retorno:

    lArqIST = False
    If pOpcao = 1 And (gM46V999.PerfilAcesso = e_Perfil_DIG_IST Or gM46V999.PerfilAcesso = e_Perfil_DIG_Yas) And Trim(pNomeArqIST) <> "" Then
        lArqIST = True
    End If

    gfImpDocumento = False
    mNomForm = pFormulario
    '1. carrego o endereço do formulário.
    Select Case UCase$(pFormulario)
    Case "MULTCOT1"
        lEndConfig = gM46V999.gEndConfCot1
        lEndImagem = gM46V999.gEndImagCot1
    Case "MULTCOT2"
        lEndConfig = gM46V999.gEndConfCot2
        lEndImagem = gM46V999.gEndImagCot2
    Case "MULTCOT3"
        lEndConfig = gM46V999.gEndConfCot3
        lEndImagem = gM46V999.gEndImagCot3
    Case "MULTPRO1"
        lEndConfig = gM46V999.gEndConfPro1
        lEndImagem = gM46V999.gEndImagPro1
    Case "MULTPRO2"
        lEndConfig = gM46V999.gEndConfPro2
        lEndImagem = gM46V999.gEndImagPro2
    Case "MULTPRO3"
        lEndConfig = gM46V999.gEndConfPro3
        lEndImagem = gM46V999.gEndImagPro3
    Case "MULTPRO4"
        lEndConfig = gM46V999.gEndConfPro4
        lEndImagem = gM46V999.gEndImagPro4
    Case "L28V009A"
        lEndConfig = gM46V999.gEndConfDeb
        lEndImagem = gM46V999.gEndImagDeb
    Case "MULTPROB"
        lEndConfig = gM46V999.gEndConfProB
        lEndImagem = gM46V999.gEndImagProB
    Case "MULTPROC"
        lEndConfig = gM46V999.gEndConfProC
        lEndImagem = gM46V999.gEndImagProC
    Case "MULTPROD"
        lEndConfig = gM46V999.gEndConfProD
        lEndImagem = gM46V999.gEndImagProD
    Case "MULTCOTA"
        lEndConfig = gM46V999.gEndConfCotA
        lEndImagem = gM46V999.gEndImagCotA
    Case "MULTCOTC"
        lEndConfig = gM46V999.gEndConfCotC
        lEndImagem = gM46V999.gEndImagCotC
    Case "ROUBOCOT1"
        lEndConfig = gM46V999.gEndConfRob1
        lEndImagem = gM46V999.gEndImagRob1
    Case "ROUBOCOT2"
        lEndConfig = gM46V999.gEndConfRob2
        lEndImagem = gM46V999.gEndImagRob2
    Case "ROUBOCOT3"
        lEndConfig = gM46V999.gEndConfRob3
        lEndImagem = gM46V999.gEndImagRob3
    Case "ROUBOCOT4"
        lEndConfig = gM46V999.gEndConfRob4
        lEndImagem = gM46V999.gEndImagRob4
    Case "ROUBOCOT5"
        lEndConfig = gM46V999.gEndConfRob5
        lEndImagem = gM46V999.gEndImagRob5
    Case "COT_LMI"
        lEndImagem = gM46V999.gEndImagCotLMI
    Case "REL_LMI"
        lEndImagem = gM46V999.gEndImagRelLMI
    Case "MULTLMIR"
        lEndImagem = gM46V999.gEndImagESPRelLMI
    Case "MULTLMIE"
        lEndImagem = gM46V999.gEndImagESPLMI
    Case Else
        Call gM46V999.gpGraLog(1, "FTE0080 - Código de formulário desconhecido: " & pFormulario & ".")
        Exit Function
    End Select

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
    If gfIsMetaframe() Then
        lArqDados = "L:\SYAS.DDD"
        mEndArq = lArqDados
        Call gM46V999.gpGraLog(0, "Eh Metaframe forcou o L " & lArqDados)
    Else
        lArqDados = gM46V999.gAppPath & "SYAS.DDD"
        mEndArq = lArqDados
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

    lListaArq = lArqDados & "|" & lEndConfig & "|" & lEndImagem & ";"

    On Error GoTo 0
    '7. Grava registros no arquivo de dados.
    Select Case pFormulario
    Case "MULTCOT1"                 'Cotação simples.
        lResultado = mfCriDados_MULTCOT1_Fonte12(lNumArquivo, pobjstcPed, pobjstcPedLoc)
    Case "MULTCOTA"
        lResultado = mfCriDados_MULTCOTA(lNumArquivo, pobjstcPed, pobjstcPedLoc)
    Case "MULTCOTC"
        lResultado = mfCriDados_MULTCOTC(lNumArquivo, pobjstcPed, pPagDe, pPagAte)
    Case "MULTCOT2"                 'Cotação coletiva - capa.
        lResultado = mfCriDados_MULTCOT2_Fonte12(lNumArquivo, pobjstcPed, pobjstcPedLoc)    'Projeto Impressão em Fonte 12
        '''            lResultado = mfCriDados_MULTCOT2(lNumArquivo, pobjstcPed, pobjstcPedLoc)
    Case "MULTCOT3"                 'Cotação coletiva - dados do item.
        lResultado = mfCriDados_MULTCOLETIVA_Itens_Fonte12(lNumArquivo, pobjstcPed, pobjstcPedLoc, pPagDe, pPagAte)    'Projeto Impressão em Fonte 12
        '''            lResultado = mfCriDados_MULTCOT3(lNumArquivo, pobjstcPed, pPagDe, pPagAte)
    Case "MULTPRO1"                 'Proposta simples - página 1.
        If lArqIST Then
            lResultado = mfCriDados_MULTPRO1(lNumArquivo, pobjstcPed, pobjstcPedLoc)
        Else
            lResultado = mfCriDados_MULTPRO1_Fonte12(lNumArquivo, pobjstcPed, pobjstcPedLoc)
        End If
    Case "MULTPRO2"                 'Proposta simples - página 2.
        lResultado = mfCriDados_MULTPRO2(lNumArquivo, pobjstcPed, pobjstcPedLoc)
    Case "MULTPRO3"                 'Proposta coletiva - capa.
        lResultado = mfCriDados_MULTPRO3_Fonte12(lNumArquivo, pobjstcPed, pobjstcPedLoc, pPagAte)    'Projeto Impressão em Fonte 12
        '''            lResultado = mfCriDados_MULTPRO3(lNumArquivo, pobjstcPed, pobjstcPedLoc, pPagAte)
    Case "MULTPRO4"                 'Proposta coletiva - dados do item.
        lResultado = mfCriDados_MULTCOLETIVA_Itens_Fonte12(lNumArquivo, pobjstcPed, pobjstcPedLoc, pPagDe, pPagAte)    'Projeto Impressão em Fonte 12
        '''            lResultado = mfCriDados_MULTPRO4(lNumArquivo, pobjstcPed, pPagDe, pPagAte)
    Case "MULTPROB"
        lResultado = mfCriDados_MULTPROB(lNumArquivo, pobjstcPed)
    Case "MULTPROC"
        lResultado = mfCriDados_MULTPROC(lNumArquivo, pobjstcPed)
    Case "MULTPROD"
        lResultado = mfCriDados_MULTPROD(lNumArquivo, pobjstcPed)
    Case "L28V009A"                 'Proposta coletiva - dados do item.
        lResultado = mfCriDados_Debito_Fonte12(lNumArquivo, pobjstcPed)
    Case "ROUBOCOT1"                'Cotação simples.
        '1000308 - SyasMulti - Roubo - Cotação Coletiva com Planos Distintos nos Itens
        lResultado = mfCriDados_ROUBOCOT1(lNumArquivo, pobjstcPed, pobjstcPedLoc, pPagDe, pPagAte)
    Case "ROUBOCOT2"
        lResultado = mfCriDados_ROUBOCOT2(lNumArquivo, pobjstcPed, pobjstcPedLoc, pPagDe, pPagAte)
    Case "ROUBOCOT3"                'Cotação coletiva - capa.
        lResultado = mfCriDados_ROUBOCOT3(lNumArquivo, pobjstcPed, pobjstcPedLoc, pPagDe, pPagAte)
    Case "ROUBOCOT4"                'Cotação coletiva - dados do item.
        lResultado = mfCriDados_ROUBOCOT4(lNumArquivo, pobjstcPed, pobjstcPedLoc, pPagDe, pPagAte)
    Case "ROUBOCOT5"
        lResultado = mfCriDados_ROUBOCOT5(lNumArquivo, pobjstcPed, pobjstcPedLoc, pPagDe, pPagAte)
    Case "COT_LMI"
        lResultado = mfCriDados_COT_LMI(lNumArquivo, pobjstcPed, lEndConfig)
    Case "REL_LMI"
        lResultado = mfCriDados_REL_LMI(lNumArquivo, pobjstcPed, lEndConfig)
    Case "MULTLMIE"
        lResultado = mfCriDados_MULTLMIE(lNumArquivo, pobjstcPed, lEndConfig)
    Case "MULTLMIR"
        lResultado = mfCriDados_MULTLMIR(lNumArquivo, pobjstcPed, lEndConfig)
    Case Else
        Call gM46V999.gpGraLog(1, " FTE0082 - Erro ao imprimir documento (código de formulário desconhecido - " & pFormulario & ").")
        Close #lNumArquivo
        Kill lArqDados
        Exit Function
    End Select

    Close #lNumArquivo

    If lVerifOverf = False Or lI = 0 Then
        lArrayArq = Split(lListaArq, ";")
        lPos = InStr(1, lArrayArq(0), "|")
        lArqDados = Left(lArrayArq(0), lPos - 1)
        lEndConfig = Mid(lArrayArq(0), lPos + 1, InStr(lPos + 1, lArrayArq(0), "|") - 1 - lPos)
        lPos = InStr(lPos + 1, lArrayArq(0), "|")
        lEndImagem = Mid(lArrayArq(0), lPos + 1)
        lCount = 1
    Else
        lPos = InStr(1, lArrayArq(lI), "|")
        lArqDados = Left(lArrayArq(lI), lPos - 1)
        lEndConfig = Mid(lArrayArq(lI), lPos + 1, InStr(lPos + 1, lArrayArq(lI), "|") - 1 - lPos)
        lPos = InStr(lPos + 1, lArrayArq(lI), "|")
        lEndImagem = Mid(lArrayArq(lI), lPos + 1)
        Call gM46V999.gpGraLog(0, lI & "  -  " & lArqDados & "  -  " & lEndConfig & "  -  " & lEndImagem)
        pImagem = Replace(pImagem, ".WMF", lI & ".WMF")
    End If

ProximoForm:

    If lResultado = False Then
        Kill lArqDados
        If pOpcao = 2 And pFormulario = "MULTCOTA" Then
            gfImpDocumento = True
        End If
        Exit Function
    End If

    'Gerar arquivo para IST
    If lArqIST Then
        lNumArquivo = FreeFile
        Open lArqDados For Input Access Read As #lNumArquivo
        lNumArqIST = FreeFile
        If Trim(Dir(pNomeArqIST)) <> "" Then
            Open pNomeArqIST For Append Shared As #lNumArqIST
        Else
            Open pNomeArqIST For Output As #lNumArqIST
        End If
        On Error GoTo 0
        Do While Not EOF(lNumArquivo)
            Line Input #lNumArquivo, lRegistro
            Print #lNumArqIST, lRegistro
        Loop
        Close #lNumArquivo
        Close #lNumArqIST
        Kill lArqDados
        'GoTo ProxArquivo
        gfImpDocumento = True
        Exit Function
    End If

    If pFormulario = "COT_LMI" Or pFormulario = "REL_LMI" Or pFormulario = "MULTLMIE" Or pFormulario = "MULTLMIR" Then
        'Antes de imprimir verifico se tem a variavel de quantidade de páginas para trocar
        Dim lNomArqAux As String
        lNumArquivo = FreeFile
        Open lArqDados For Input Access Read As #lNumArquivo
        lNumArqIST = FreeFile
        lNomArqAux = Replace(lArqDados, ".DDD", ".DD1")
        If Dir(lNomArqAux) <> "" Then
            On Error Resume Next
            Kill lNomArqAux
            On Error GoTo 0
        End If
        If Dir(lNomArqAux) <> "" Then
            lNomArqAux = Replace(lArqDados, ".DDD", ".DD2")
        End If
        If Dir(lNomArqAux) <> "" Then
            On Error Resume Next
            Kill lNomArqAux
            On Error GoTo 0
        End If
        If Dir(lNomArqAux) <> "" Then
            lNomArqAux = Replace(lArqDados, ".DDD", ".DD3")
        End If

        Open lNomArqAux For Append Shared As #lNumArqIST
        On Error GoTo 0
        Do While Not EOF(lNumArquivo)
            Line Input #lNumArquivo, lRegistro
            lRegistro = Replace(lRegistro, "@@TOTALPAG@@", lTotPagina)
            Print #lNumArqIST, lRegistro
        Loop

        Close #lNumArquivo
        Close #lNumArqIST
        lArqDados = lNomArqAux
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
    DoEvents
    Err = 0
    Call PrintReld(CStr(pOpcao))        '<<<<<<<<<<========== Comando que apresentou erro - Dr. Watson...
    If Err <> 0 Then
        'para o form "COT_LMI"esta dando esse erro 6
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
        gfImpDocumento = True
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

ProxArquivo:
    For lI = lCount To UBound(lArrayArq) - 1

        lPos = InStr(1, lArrayArq(lI), "|")
        lArqDados = Left(lArrayArq(lI), lPos - 1)
        lEndConfig = Mid(lArrayArq(lI), lPos + 1, InStr(lPos + 1, lArrayArq(lI), "|") - 1 - lPos)
        lPos = InStr(lPos + 1, lArrayArq(lI), "|")
        lEndImagem = Mid(lArrayArq(lI), lPos + 1)
        pImagem = Replace(pImagem, ".WMF", lI & ".WMF")
        lCount = lCount + 1
        If gM46V999.gfPreenchido(lEndConfig) And Dir(lEndConfig) <> "" Then
            GoTo ProximoForm
        Else
            Exit For
        End If
    Next lI
    gfImpDocumento = True
End Function


Public Sub mpGerarEspecificacaoLMI(ByVal pNum_Cotac_LMI As String, ByVal pNumPI As Long, ByVal pCodRamo As Long)

    Dim lAuxiliar As String     'Auxiliar.
    Dim lImagem As String   'Nome para arquivo imagem.
    Dim lMensagem As String     'Mensagem.
    Dim lobjstcLoc As stcA46V708B    'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
    Dim lCount As Integer    'Contador.
    Dim lobjstcCob As stcA46V704B
    Dim colecaoTMP As Collection
    Dim colecaoTMP2 As Collection
    Dim lCoberturas As String
    Dim lbdTAB_Cobert As ADODB.Connection
    Dim lrsTab_Cobert As clsYasRecordSet
    Dim lI As Integer
    Dim lArray As Variant
    Dim lObjPed As clsA46V702A    'Objeto da classe clsA46V702A (TAB_PED).
    Dim lObjstcPed As stcA46V702B    'Objeto da classe clsA46V702A (TAB_PED).
    Dim lOpcao As Byte
    Dim lAux As String
    Dim lNomeArq As String
    Dim lSelect As String
    Dim lrsAux As clsYasRecordSet
    Dim DtEntrada As String
    Dim Ano As String
    Dim Mes As String
    Dim Dia As String
    Dim Caminho As String
    Dim CaminhoArq As String
    Dim lAplicacao As Word.Application
    Dim lDocumento As Object
    Dim lDigito As String

    Dim clsA99V001 As clsA99V001
    Dim lInsert As String

    On Error GoTo ControleErro

    Set lObjPed = New clsA46V702A
    Set lObjPed.M46V999 = gM46V999

    Set lObjstcPed = New stcA46V702B
    Set lObjstcPed.M46V999 = gM46V999



    lObjstcPed.NOSSO_NUMERO = pNum_Cotac_LMI

    If Not lObjPed.mfCarregarColecao(lObjstcPed, "", "", True) Then
        Call gM46V999.gpGraLog(1, "GNP0024 - " & lObjPed.UltimoErro)
        Exit Sub
    End If

    If lObjPed.Colecao.Count <> 1 Then
        Call gM46V999.gpGraLog(1, "GNP0025 - Cotação de Referencia não existe!")
        Exit Sub
    End If

    Set lObjstcPed = lObjPed.Colecao.Item(1)

    If lObjstcPed.LMI_Unico <> "1" Then
        Call gM46V999.gpGraLog(1, "GNP0026 - Cotação de Referencia não é Cotacao de LMI!")
        Exit Sub
    End If

    'Primeiramente, as coberturas serão ordenadas.
    Set colecaoTMP2 = New Collection
    For Each lobjstcLoc In lObjstcPed.ColecaoPedLoc
        Set colecaoTMP = New Collection
        lCoberturas = ""
        For Each lobjstcCob In lobjstcLoc.ColecaoPedCobert
            If gM46V999.gfPreenchido(lCoberturas) Then
                lCoberturas = lCoberturas & ", "
            End If
            lCoberturas = lCoberturas & lobjstcCob.Cod_Cobert
        Next lobjstcCob


        If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTAB_Cobert, lMensagem) = False Then
            Call gM46V999.gpGraLog(2, lMensagem)
            Exit Sub
        End If

        lSelect = " Select * from Tab_Cobert where cod_cobert in (" & lCoberturas & ")"
        lSelect = lSelect & " and cod_Ramo = " & gM46V999.gCod_Ramo
        If val(Right(frmT46V101A.txtRisRegiao.Tag, 4)) <> 0 And gM46V999.gCodProduto <= gM46V999.TarifaAntiga Then
            lSelect = lSelect & " and tab_cobert = " & val(Right(frmT46V101A.txtRisRegiao.Tag, 4))
        Else
            lSelect = lSelect & " and tab_cobert = " & frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex)
        End If
        lSelect = lSelect & " and COD_TARIF_INI <= " & gM46V999.gCodProduto
        lSelect = lSelect & " and COD_TARIF_TER >= " & gM46V999.gCodProduto
        lSelect = lSelect & " order by NUM_ORDEM_COBERT"

        If gM46V999.gfObtRegistro(lbdTAB_Cobert, lSelect, lrsTab_Cobert, lMensagem) = False Then
            Call gM46V999.gpFecha2(lbdTAB_Cobert)
            Call gM46V999.gpGraLog(2, lMensagem)
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        lCoberturas = ""
        Do While Not lrsTab_Cobert.EOF
            lCoberturas = lCoberturas & Format(lrsTab_Cobert("Cod_Cobert"), "00") & ";"
            lrsTab_Cobert.MoveNext
        Loop

        lArray = Split(lCoberturas, ";")
        For lI = 0 To UBound(lArray)
            If gM46V999.gfPreenchido(lArray(lI)) Then
                Set lobjstcCob = lobjstcLoc.ColecaoPedCobert.Item(lobjstcLoc.Class_ID & Format(lArray(lI), "0000") & Format(gM46V111.gCod_plano, "00000"))
                colecaoTMP.Add lobjstcCob, lobjstcCob.Class_ID
            End If
        Next lI
        Set lobjstcLoc.ColecaoPedCobert = Nothing
        Set lobjstcLoc.ColecaoPedCobert = colecaoTMP
        colecaoTMP2.Add lobjstcLoc, lobjstcLoc.Class_ID
    Next lobjstcLoc

    Set lObjstcPed.ColecaoPedLoc = Nothing
    Set lObjstcPed.ColecaoPedLoc = colecaoTMP2

    '3. Imprime documento.
    lOpcao = 2    'Gerar Imagem

    lImagem = "LMI_COT" & Format$(lObjstcPed.Num_Ped, "000000") & ".WMF"
    lAuxiliar = lImagem
    If gfImpDocumento(lOpcao, "MULTLMIE", lAuxiliar, lObjstcPed, lobjstcLoc, 1, 1, lNomeArq) = False Then
        Exit Sub
    End If
    lAuxiliar = "LMI_COT" & Format$(lObjstcPed.Num_Ped, "000000") & "_REL.WMF"
    If gfImpDocumento(lOpcao, "MULTLMIR", lAuxiliar, lObjstcPed, lobjstcLoc, 1, 1, lNomeArq) = False Then
        Exit Sub
    End If

    'Salvo em Word
    Call gpGeraFiguraArqWord(lImagem, 0, 1)

    lAuxiliar = gM46V999.gAppPath & Replace(lImagem, "WMF", "DOC")
    lNomeArq = Format(pNumPI, "0000000000") & "_ES02.DOC"
    If Dir(gM46V999.gAppPath & lNomeArq) <> "" Then
        Kill gM46V999.gAppPath & lNomeArq
    End If
    Name gM46V999.gAppPath & Replace(lImagem, "WMF", "DOC") As gM46V999.gAppPath & lNomeArq
    lAuxiliar = gM46V999.gAppPath & lNomeArq

    If UCase(gM46V999.gSerDIVERSOS) Like "*MIX*" Then
        gM46V999.gPathImagePI = gM46V999.gAppPath
    Else
        gM46V999.gPathImagePI = "\\SPX21111PSLUKFL\proposta4\"   'Thais solicitou alterar diretorio p/ (proposta4) - 05/06/2012
    End If
    lSelect = "Select * from ged.dbo.tab_image_Proposta where num_pi = " & pNumPI & " order by dat_digitalizacao"
    '
    Dim lbdBasDados As ADODB.Connection    'Base de dados.'''

    If gM46V999.gfAbrBasDados("GED", gM46V999.gTipGed, gM46V999.gEndGed, lbdBasDados, lMensagem) = False Then
        Call gM46V999.gpGraLog(1, lMensagem)
        Exit Sub
    End If

    If gM46V999.gfObtRegistro(lbdBasDados, lSelect, lrsAux, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If

    If Not lrsAux.EOF Then
        DtEntrada = Format(Trim(lrsAux("dat_digitalizacao")), "dd/mm/yyyy")
        gM46V999.gPathImagePI = lrsAux("dir_img")
        Caminho = lrsAux("dir_img")    'incluído 25/11/2011 - por Marcos e Thais
    Else
        DtEntrada = Format(Now, "dd/mm/yyyy")
        Ano = Right(DtEntrada, 4)
        Mes = Mid(DtEntrada, 4, 2)
        Dia = Left(DtEntrada, 2)

        Caminho = gM46V999.gPathImagePI & Ano
        If Dir(Caminho, vbDirectory) = "" Then
            MkDir Caminho
        End If
        Caminho = Caminho & "\" & Mes
        If Dir(Caminho, vbDirectory) = "" Then
            MkDir Caminho
        End If
        Caminho = Caminho & "\" & Dia
        If Dir(Caminho, vbDirectory) = "" Then
            MkDir Caminho
        End If
        Caminho = Caminho & "\" & Format(pNumPI, "0000000000") & "\"
        If Dir(Caminho, vbDirectory) = "" Then
            MkDir Caminho
        End If
    End If

    CaminhoArq = "S:\Especificacao\"
    'lNomeArq = "ESP_" & Format(pNumPI, "0000000000") & "_" & Format(lCount, "0000") & ".DOC"
    'lNomeArq = "ESP_" & Format(pNumPI, "0000000000") & ".DOC"

    Dim clsPDF As Object
    Dim clsOpcao As Object

    Set lAplicacao = CreateObject("Word.Application")
    On Error Resume Next
AbreNovamente:
    lAplicacao.Documents.Open lAuxiliar
    If Err Then
        'essa coisa feia abaixo é para funcionar no Office 2010, abrindo o doc novamente não da erro
        If Err.Number = 5132 Then
            Err.Clear
            GoTo AbreNovamente
        Else
            GoTo ControleErro
        End If
    End If
    lAplicacao.WordBasic.FilePrintSetup Printer:="PDFCreator", DoNotSetAsSysDefault:=1

    lAplicacao.PrintOut Range:=0, Item:= _
                        0, Copies:=1, Pages:="", PageType:=0, _
                        Collate:=True, Background:=True, PrintToFile:=False

    '    lAplicacao.PrintOut Range:=wdPrintAllDocument, Item:= _
         '        wdPrintDocumentContent, Copies:=1, Pages:="", PageType:=wdPrintAllPages, _
         '        Collate:=True, Background:=True, PrintToFile:=False

    Call gM46V999.gpGraLog(1, "Será exibida uma janela do WORD com o questionamento sobre as margens de impressão. Por favor clicar SIM e depois em OK desta mensagem.")

    lAplicacao.Documents.Close
    lAplicacao.Quit

    Set lAplicacao = Nothing
    On Error GoTo ControleErro

    lAux = CaminhoArq & lNomeArq

    If Dir(Replace(UCase(lAuxiliar), ".DOC", ".PDF")) <> "" Then
        If Not mfFileCopy(Replace(UCase(lAuxiliar), ".DOC", ".PDF"), Replace(UCase(lAux), ".DOC", ".PDF")) Then
            Exit Sub
        End If
    End If


    lSelect = "Select * from ged.dbo.tab_image_proposta where num_pi = " & pNumPI
    lSelect = lSelect & " and nom_img = '" & Replace(lNomeArq, ".DOC", ".PDF") & "'"

    If gM46V999.gfObtRegistro(lbdBasDados, lSelect, lrsAux, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If

    If lrsAux.EOF Then
        lSelect = "Select max(id_doc) + 1 as contador from ged.dbo.tab_image_proposta where num_pi = " & pNumPI

        If gM46V999.gfObtRegistro(lbdBasDados, lSelect, lrsAux, lMensagem) = False Then
            Call gM46V999.gpGraLog(2, lMensagem)
            Exit Sub
        End If

        lCount = 1

        If Not lrsAux.EOF Then
            If Trim(lrsAux("contador")) <> "" Then
                lCount = Trim(lrsAux("contador"))
            End If
        End If

        lInsert = "insert into ged.dbo.tab_image_proposta (num_pi,id_doc,nom_img,dir_img,dat_digitalizacao,"
        lInsert = lInsert & "nom_membro) values (" & pNumPI & "," & lCount & ",'" & Replace(lNomeArq, ".DOC", ".PDF") & "'"
        lInsert = lInsert & ",'" & Caminho & "',convert(datetime,'" & Format(Now(), "dd/mm/yyyy hh:mm:ss.000") & "',103),'SYASMUL')"

        If gM46V999.gfExeSQL(lbdBasDados, lInsert, lAux) <> 0 Then
            Call gM46V999.gpFecha2(lbdTAB_Cobert)
            Call gM46V999.gpGraLog(2, lMensagem)
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        lSelect = "Select * from ged.dbo.tab_ctrl_espec where num_pi = " & pNumPI
        If gM46V999.gfObtRegistro(lbdBasDados, lSelect, lrsAux, lMensagem) = False Then
            Call gM46V999.gpGraLog(2, lMensagem)
            Exit Sub
        End If

        If lrsAux.EOF Then
            lInsert = "insert into ged.dbo.tab_ctrl_espec (Cod_Ramo, Num_PI, NomeArq, Dat_Inclusao, Dat_Anexo)"
            lInsert = lInsert & " values (" & pCodRamo & "," & pNumPI & ",'" & Mid(lNomeArq, 1, 20) & "',"
            lInsert = lInsert & Format(Now(), "yyyymmdd") & "," & Format(Now(), "yyyymmdd") & ")"

            If gM46V999.gfExeSQL(lbdBasDados, lInsert, lAux) <> 0 Then
                Call gM46V999.gpFecha2(lbdTAB_Cobert)
                Call gM46V999.gpGraLog(2, lAux)
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
    End If

    On Error Resume Next
    Kill lAuxiliar
    Screen.MousePointer = vbDefault
    Exit Sub

ControleErro:
    Call gM46V999.gpGraLog(1, "mpGerarEspecificacaoLMI ERRO - " & Err & " - " & Error)
    Screen.MousePointer = vbNormal
    If Not lAplicacao Is Nothing Then
        If lAplicacao.Documents.Count > 0 Then
            lAplicacao.Documents.Close
        End If
        lAplicacao.Quit
    End If
    Exit Sub
    Resume Next
End Sub




Private Function mfCriDados_MULTCOT1(ByVal pNumArqDados As Integer, ByVal pobjstcPed As stcA46V702B, _
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
    Dim lPrazo As Integer
    Dim lCodCobert As Integer
    Dim lNumMeses As Integer
    Dim lDescCom As Double
    Dim lDescExp As Double
    Dim lValorComis As Single
    Dim lbdTbm_Regiao As ADODB.Connection
    Dim lrsTbm_Regiao As clsYasRecordSet
    Dim lQuaDias As Integer
    Dim lSelect As String
    mfCriDados_MULTCOT1 = False

    '001. Título do seguro (de acordo com a base de dados).
    '     Neste caso deve ser colocado YASUDA RESIDENCIAL, YASUDA CONDOMÍNIO ou YASUDA EMPRESARIAL.
    lAuxiliar2 = " YASUDA " & UCase(pobjstcPed.gDesc_Ramo)

    Do While Len(lAuxiliar2) < 36
        lAuxiliar2 = " " & lAuxiliar2 & " "
    Loop

    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)


    '[FRH] - FICHA 901116
    '001B. Código do Processo Susep.
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


    '002. Emissão.
    Print #pNumArqDados, "Emissão: " & Format$(Now, "dd/mm/yyyy hh:nn:ss")

    '003. Validade.
    If pobjstcPedLoc.Tip_Emissao = e_TipEmissao_RenYasuda Then
        If gM46V999.gCodProduto >= 978 Then
            Print #pNumArqDados, "Validade: " & Format$(CDate(gM46V999.gfForData(2, pobjstcPed.Dat_Pri_Calculo)) + 15, "dd/mm/yyyy")
        Else
            Print #pNumArqDados, "Validade: " & Format$(CDate(gM46V999.gfForData(2, pobjstcPed.Dat_Pri_Calculo)) + 30, "dd/mm/yyyy")
        End If
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
    If pobjstcPed.ColecaoPedCorr.Count > 0 Then
        Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(1)
        lAuxiliar1 = lObjstcCorr.Cod_Corr
        lAuxiliar2 = lObjstcCorr.gNom_Corretor
    Else
        lAuxiliar1 = 0
        lAuxiliar2 = " "
    End If

    Print #pNumArqDados, Format$(lAuxiliar1, "0000") & " - " & lAuxiliar2

    '007. Tarifa .
    lAuxiliar2 = pobjstcPed.Cod_Produto
    If pobjstcPed.Cod_Produto = 976 And pobjstcPed.Cod_Ramo = e_CodRamo_Empresarial Then
        lAuxiliar2 = lAuxiliar2 & " A"
    End If
    If pobjstcPed.Cod_Sequencia = 1 Then
        lAuxiliar2 = lAuxiliar2 & " (A)"
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '008. Nosso número .
    lAuxiliar2 = pobjstcPed.NOSSO_NUMERO
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)


    If pobjstcPed.objstcCtrlEmis.Tip_Cotac = 2 And pobjstcPed.objstcCtrlEmis.Num_Cotac <> 0 Then
        Print #pNumArqDados, "Siscota: " & pobjstcPed.objstcCtrlEmis.Num_Cotac
    Else
        Print #pNumArqDados, " "
    End If
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
        'lAuxiliar2 = "CNPJ.........:"
        lAuxiliar2 = "              "
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        lAuxiliar2 = pobjstcPed.gCNPJ_CPF_Formatado & "   " & _
                     "Ramo de atividade: " & pobjstcPed.gDesc_RamoAtiv
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    Case e_TipPessoa_Fisica
        Print #pNumArqDados, "Física"
        '012 e 013. Proponente - CNPJ ou CPF .
        'lAuxiliar2 = "CPF...........:"
        lAuxiliar2 = "              "
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
    'Ficha 900417
    Do While Len(lAuxiliar2) < 70
        lAuxiliar2 = lAuxiliar2 & " "
    Loop

    If gM46V999.gCodProduto >= 968 Then
        lAuxiliar2 = lAuxiliar2 & "Profissão: " & pobjstcPed.Nom_Empresa
    Else
        lAuxiliar2 = lAuxiliar2 & "Empresa: " & pobjstcPed.Nom_Empresa
    End If
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
    If gM46V999.gCod_Ramo = e_CodRamo_Empresarial Then
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
    If pobjstcPed.Cod_Produto <= gM46V999.TarifaAntiga Then
        lAuxiliar2 = pobjstcPedLoc.Cod_Regiao & " - " & Left(pobjstcPedLoc.gDesc_Regiao, 38)
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    Else
        'Nova Região
        If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTbm_Regiao, lMensagem) = False Then
            Call gM46V999.gpGraLog(2, lMensagem)
            Exit Function
        End If

        lSelect = "Select Cod_Faixa_Reg, DSC_REG from Tbm_Faixa_Regiao "
        lSelect = lSelect & " Where cod_Ramo = " & pobjstcPed.Cod_Ramo
        lSelect = lSelect & "    and Cod_Modalidade = " & pobjstcPedLoc.Cod_Produto
        lSelect = lSelect & "    and Cod_Tarif = " & pobjstcPed.Cod_Produto
        lSelect = lSelect & "    and COD_Sequencia <= " & pobjstcPed.Cod_Sequencia
        lSelect = lSelect & "    and Cep_Ini <= " & pobjstcPedLoc.Num_Cep_Risco
        lSelect = lSelect & "    and Cep_Fim >= " & pobjstcPedLoc.Num_Cep_Risco
        lSelect = lSelect & "    order by Cod_Sequencia desc "

        If gM46V999.gfObtRegistro(lbdTbm_Regiao, lSelect, lrsTbm_Regiao, lMensagem) = False Then
            Call gM46V999.gpGraLog(0, lMensagem)
            Call gM46V999.gpFecha2(lbdTbm_Regiao)
            Exit Function
        End If

        If Not lrsTbm_Regiao.EOF Then
            lAuxiliar2 = gM46V999.Check(lrsTbm_Regiao("Cod_Faixa_Reg")) & " - " & Left(gM46V999.Check(lrsTbm_Regiao("DSC_REG")), 38)
        Else
            lAuxiliar2 = "Não cadastrada."
        End If
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    End If

    '032. Dados do seguro - vigência .

    lAuxiliar2 = "de " & gM46V999.gfForData(2, pobjstcPed.Dat_Ini_Vig) & _
               " a " & gM46V999.gfForData(2, pobjstcPed.Dat_Ter_Vig)

    'Se for perfil Calculo não será impresso conforme ficha 0701831
    If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda And _
       (gM46V999.PerfilAcesso = e_Perfil_CALC Or gM46V999.PerfilAcesso = e_Perfil_CALC_Gestor) Then
        lQuaDias = gM46V999.gfDifDatas(pobjstcPed.Dat_Ini_Vig, pobjstcPed.Dat_Ter_Vig)
        If lQuaDias >= 365 Then
            lAuxiliar2 = ""
        End If
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '033. Dados do seguro - valor em risco .
    lAuxiliar2 = "R$ " & Format(pobjstcPedLoc.Val_Vr, "standard")
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '034. Dados do seguro - tipo de emissão .
    lAuxiliar2 = pobjstcPedLoc.gDesc_TipEmissao
    Select Case pobjstcPedLoc.Tip_Emissao
    Case e_TipEmissao_RenYasuda, e_TipEmissao_Endosso
        lAuxiliar2 = lAuxiliar2 & " " & "  Apólice: " & pobjstcPedLoc.Num_Apol_Anterior & "  Item: " & pobjstcPedLoc.Num_Item_Anterior & "  Sinistro: " & IIf(pobjstcPedLoc.Ind_Sinistro = e_SimNao_Sim, "Sim", "Não")
    End Select
    Print #pNumArqDados, lAuxiliar2

    'm If pObjstcPed.Cod_Ramo = e_CodRamo_Empresarial Then
    '035 e 036. Coberturas contratadas - título.
    'm lAuxiliar2 = "Desconto de"
    'm Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    'm lAuxiliar2 = "experiência"
    'm Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    'm Else
    Print #pNumArqDados, " "
    Print #pNumArqDados, " "
    'm End If

    'daqui prá baixo

    '037 a 212 Coberturas contratadas - linhas de 01 a 22.
    lCountLine = 0
    lValTotCober = 0
    For Each lobjstcCob In pobjstcPedLoc.ColecaoPedCobert
        lCountLine = lCountLine + 1
        'Código
        Print #pNumArqDados, Format$(lobjstcCob.Cod_Cobert, "0000")
        lCodCobert = lobjstcCob.Cod_Cobert
        'Descrição
        Print #pNumArqDados, Left(lobjstcCob.gNomeCobertura, 45)

        lAuxiliar1 = IIf(IsNumeric(lobjstcCob.Num_Meses_PI), lobjstcCob.Num_Meses_PI, 0)
        If lAuxiliar1 = 0 Then
            Print #pNumArqDados, " "

            Print #pNumArqDados, " "
        Else
            'pi mes
            Print #pNumArqDados, Format$(lobjstcCob.Num_Meses_PI, "00")
            lNumMeses = lobjstcCob.Num_Meses_PI
            Call mpPrazoTabCoefPI(pobjstcPed, lCodCobert, lNumMeses, lPrazo)
            'pi dias
            If lPrazo <> 0 Then
                Print #pNumArqDados, lPrazo
            Else
                Print #pNumArqDados, " "
            End If
        End If
        '%Sobre Prejuízos
        lAuxiliar1 = lobjstcCob.Per_Franq
        lAuxiliar2 = Format$(lAuxiliar1, "standard")
        lTamanho = Len(lAuxiliar2)
        If lTamanho < 10 Then
            lAuxiliar2 = Space(10 - lTamanho) & lAuxiliar2
        End If
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        'Limitado ao Mínimo
        lAuxiliar1 = lobjstcCob.Val_Franq_Min
        lAuxiliar2 = Format$(lAuxiliar1, "standard")
        lTamanho = Len(lAuxiliar2)
        If lTamanho < 14 Then
            lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
        End If
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

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
        lDescCom = lobjstcCob.Per_Desc_Coml
        lAuxiliar2 = Format$(lAuxiliar1, "standard")
        lTamanho = Len(lAuxiliar2)
        If lTamanho < 10 Then
            lAuxiliar2 = Space(10 - lTamanho) & lAuxiliar2
        End If
        'm        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        If pobjstcPed.Cod_Ramo = e_CodRamo_Empresarial Then
            'e) Desconto de experiência .
            lAuxiliar1 = lobjstcCob.Per_Desc_Exp
            lDescExp = lobjstcCob.Per_Desc_Exp
            lAuxiliar2 = Format$(lAuxiliar1, "standard")
            lTamanho = Len(lAuxiliar2)
            If lTamanho < 10 Then
                lAuxiliar2 = Space(10 - lTamanho) & lAuxiliar2
            End If
            'm            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        Else
            lDescExp = 0
            'm            Print #pNumArqDados, " "
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
        Print #pNumArqDados, " "
        Print #pNumArqDados, " "
    Loop

    Print #pNumArqDados, "              Soma:"
    '213. Coberturas contratadas - soma.
    lAuxiliar1 = lValTotCober
    lAuxiliar2 = Format$(lAuxiliar1, "standard")
    lTamanho = Len(lAuxiliar2)
    If lTamanho < 14 Then
        lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    'Calcular Soma da Comissão para saber se é premio NET
    If pobjstcPed.ColecaoPedCorr.Count > 0 Then
        Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(1)
        lValorComis = lObjstcCorr.Per_Corr
        If pobjstcPed.ColecaoPedCorr.Count > 1 Then
            Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(2)
            lValorComis = lValorComis + lObjstcCorr.Per_Corr
            If pobjstcPed.ColecaoPedCorr.Count > 2 Then
                Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(3)
                lValorComis = lValorComis + lObjstcCorr.Per_Corr
            End If
        End If
    End If

    If lValorComis = 0 Then
        Print #pNumArqDados, "        Prêmio NET:"
    Else
        Print #pNumArqDados, "            Prêmio:"
    End If

    '214. Coberturas contratadas - prêmio líquido.
    lAuxiliar1 = pobjstcPed.VAL_PR
    lAuxiliar2 = Format$(lAuxiliar1, "standard")
    lTamanho = Len(lAuxiliar2)
    If lTamanho < 14 Then
        lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '*************** Premio NET
    If lValorComis = 0 Then
        For lI = 1 To 6
            Print #pNumArqDados, " "
        Next
        GoTo ImprimirLegendas
    End If
    '*************** Premio NET
    Print #pNumArqDados, "  Custo da Apólice:"
    '215. Coberturas contratadas - custo da apólice.
    lAuxiliar1 = pobjstcPed.Val_Cust_Emis
    lAuxiliar2 = Format$(lAuxiliar1, "standard")
    lTamanho = Len(lAuxiliar2)
    If lTamanho < 14 Then
        lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    lAuxiliar2 = "IOF (" & Format(gM46V111.gProJurIOF, "standard") & "%):"
    Print #pNumArqDados, String(19 - Len(lAuxiliar2), " ") & lAuxiliar2

    '216. Coberturas contratadas - IOF.
    lAuxiliar1 = pobjstcPed.Val_IOF
    lAuxiliar2 = Format$(lAuxiliar1, "standard")
    lTamanho = Len(lAuxiliar2)
    If lTamanho < 14 Then
        lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    Print #pNumArqDados, "Prêmio total vista:"
    '217. Coberturas contratadas - prêmio total a vista.
    lAuxiliar1 = pobjstcPed.gValorTotal
    lAuxiliar2 = Format$(lAuxiliar1, "standard")
    lTamanho = Len(lAuxiliar2)
    If lTamanho < 14 Then
        lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    'Premio NET
ImprimirLegendas:

    '218. Informação cifrada .

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
    '2111 a 2120 - Embraer
    'Ficha 901772
    '2121 a 2130 - Bardella

    If (gM46V111.gCod_plano >= 2001 And gM46V111.gCod_plano <= 2049) Or (gM46V111.gCod_plano >= 2051 And gM46V111.gCod_plano <= 2100) Or (gM46V111.gCod_plano >= 2111 And gM46V111.gCod_plano <= 2130) Then
        lAuxiliar2 = "B:00"
    Else
        lAuxiliar2 = "B:" & Format(pobjstcPedLoc.Per_Bonus, "00")
    End If

    '* lAuxiliar2 = lAuxiliar2 & "/A:" & Format(pobjstcPedLoc.Per_Desc_Apar, "00")

    'Ficha 900740
    If gM46V111.gCod_plano >= 2021 And gM46V111.gCod_plano <= 2031 Then
        lAuxiliar2 = lAuxiliar2 & "/P:00"
    Else
        lAuxiliar2 = lAuxiliar2 & "/P:" & Format(pobjstcPedLoc.Per_Desc_Progres, "00")
    End If
    'lAuxiliar2 = lAuxiliar2 & "/CO:" & Format(pObjstcPed.objstcCtrlEmis.Per_Corr + pObjstcPed.objstcCtrlEmis.Per_Corr_2 + pObjstcPed.objstcCtrlEmis.Per_Corr_3, "00.00")
    lAuxiliar1 = 0

    If pobjstcPed.ColecaoPedCorr.Count > 0 Then
        Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(1)
        lAuxiliar1 = lObjstcCorr.Per_Corr
        If pobjstcPed.ColecaoPedCorr.Count > 1 Then
            Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(2)
            lAuxiliar1 = lAuxiliar1 + lObjstcCorr.Per_Corr
            If pobjstcPed.ColecaoPedCorr.Count > 2 Then
                Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(3)
                lAuxiliar1 = lAuxiliar1 + lObjstcCorr.Per_Corr
            End If
        End If
    End If

    lAuxiliar2 = lAuxiliar2 & "/CO:" & Format(lAuxiliar1 * 100, "0000")
    '* lAuxiliar2 = lAuxiliar2 & "/C:" & Format(Round(pobjstcPedLoc.Val_Coef_IS_Vr, 4) * 10000, "00000")

    'Ficha 900740
    If gM46V111.gCod_plano >= 2021 And gM46V111.gCod_plano <= 2031 Then
        lAuxiliar2 = lAuxiliar2 & "/DC:00"
    Else
        lAuxiliar2 = lAuxiliar2 & "/DC:" & Format(lDescCom * 100, "0000")
    End If
    '* lAuxiliar2 = lAuxiliar2 & "/DE:" & Format(lDescExp, "00")
    'lAuxiliar2 = lAuxiliar2 & "/PRGRS6"

    'Ficha 900550
    '/V1.0 -> versão após ajustar o bônus das coberturas
    lAuxiliar2 = lAuxiliar2 & "/V1.0"

    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)


    '*************** Premio NET
    If lValorComis = 0 Then
        For lI = 1 To 24
            Print #pNumArqDados, " "
        Next
        GoTo ImprimirFranquias
    End If
    '*************** Premio NET


    'Calcula o fracionamento
    With pobjstcPed
        Call gM46V998.gfCalFracionamento(pobjstcPed, pobjstcPedLoc.Tip_Emissao, _
                                         lValorParc, lMensagem, gM46V111.gProJurIOF, .Cod_Reg, .Val_IOF, .Val_Parc_Min)
    End With
    '219 a 230. Formas de pagamento - primeira parcela a vista.

    'Ficha 900665
    For lI = 1 To 12
        If lI > pobjstcPed.Qtd_Parc Then
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
                    If lTamanho < 9 Then
                        lDemParcelas = Space(9 - lTamanho) & lDemParcelas
                    End If
                    'd) Juros.
                    Set lobjstcJur = New stcA46V711B
                    On Error Resume Next
                    Set lobjstcJur = pobjstcPed.ColecaoPedJuros.Item(pobjstcPed.NOSSO_NUMERO & Format(lI, "00"))
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

    '231 a 242. Formas de pagamento - primeira parcela a 30 dias.

    'Ficha 900665
    For lI = 1 To 12
        If lI > pobjstcPed.Qtd_Parc Then
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
                    If lTamanho < 10 Then
                        lPriParcela = Space(10 - lTamanho) & lPriParcela
                    End If
                    'c) Valor das demais parcelas.
                    lAuxiliar1 = lValorParc(lI, 4)
                    lDemParcelas = Format$(lAuxiliar1, "standard")
                    lTamanho = Len(lDemParcelas)
                    If lTamanho < 9 Then
                        lDemParcelas = Space(9 - lTamanho) & lDemParcelas
                    End If
                    'd) Juros.
                    Set lobjstcJur = New stcA46V711B
                    On Error Resume Next
                    Set lobjstcJur = pobjstcPed.ColecaoPedJuros.Item(pobjstcPed.NOSSO_NUMERO & Format(lI, "00"))
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
    'Premio NET
ImprimirFranquias:

    '243 a 250. Franquias.
    lCountLine = 0
    Print #pNumArqDados, "FRANQUIAS"
    lCountLine = lCountLine + 1
    'Call mpDescricaoFranquia_Corretor(pobjstcPedLoc, pobjstcPed, lDescFranquias, 28 - lCountLine)

    '
    'Não deve mais buscar pois já está gravado
    'Call gM46V999.mpNovaBuscaDescricaoFranquia(pobjstcPed, pobjstcPedLoc, 28 - lCountLine, 0, lDescFranquias)

    For lI = 0 To 27 - lCountLine
        If lDescFranquias(lI) = "" Then Exit For
        lAuxiliar2 = lDescFranquias(lI)
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        lCountLine = lCountLine + 1
    Next lI

    'Call mpDescricaoFranquia(pobjstcPedLoc, pobjstcPed, lDescFranquias, 28 - lCountLine)
    'For li = 0 To 27 - lCountLine
    '    If lDescFranquias(li) = "" Then Exit For
    '    lAuxiliar2 = lDescFranquias(li)
    '    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    '    lCountLine = lCountLine + 1
    'Next li

    '251 a 260. Importante.
    Print #pNumArqDados, " "
    Print #pNumArqDados, "IMPORTANTE"
    lCountLine = lCountLine + 2
    Call mpTextoImportante(pobjstcPedLoc, pobjstcPed, lDescImportante, "MULTCOT1", 28 - lCountLine)
    For lI = 0 To UBound(lDescImportante) - 1
        If lDescImportante(lI) = "" Then Exit For
        lAuxiliar2 = lDescImportante(lI)
        If gM46V999.gfPreenchido(lAuxiliar2) And lCountLine < 28 Then
            lCountLine = lCountLine + 1
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        End If
    Next lI
    Call mpTextoObs(pobjstcPedLoc, lDescImportante, 28 - lCountLine)
    For lI = 0 To UBound(lDescImportante) - 1
        If lDescImportante(lI) = "" Then Exit For
        lAuxiliar2 = lDescImportante(lI)
        If gM46V999.gfPreenchido(lAuxiliar2) Then
            lCountLine = lCountLine + 1
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        End If
    Next lI

    If gM46V999.gCodProduto >= 958 And pobjstcPedLoc.Cob_Bens = True And Not mfVerDigBens() Then
        For Each lobjstcCob In pobjstcPedLoc.ColecaoPedCobert
            If gM46V999.gCod_Ramo = e_CodRamo_Empresarial And lobjstcCob.Cod_Cobert = gM46V999.gfBuscarDePara(94, 1) And lCountLine < 28 Then
                lAuxiliar2 = "Para a cobertura " & lobjstcCob.Cod_Cobert & " é obrigatório relacionar os bens e respectivos valores. A relação poderá ser preenchida diretamente no SYAS"
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
                lCountLine = lCountLine + 1
                lAuxiliar2 = "ou ser entregue juntamente com a proposta de seguro."
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
                lCountLine = lCountLine + 1
            End If
            If gM46V999.gCod_Ramo = e_CodRamo_Residencial And (lobjstcCob.Cod_Cobert = gM46V999.gfBuscarDePara(27, 1) Or _
                                                               lobjstcCob.Cod_Cobert = gM46V999.gfBuscarDePara(29, 1) Or _
                                                               lobjstcCob.Cod_Cobert = gM46V999.gfBuscarDePara(30, 1) Or _
                                                               lobjstcCob.Cod_Cobert = gM46V999.gfBuscarDePara(31, 1)) And lCountLine < 28 Then
                lAuxiliar2 = "Para a cobertura " & lobjstcCob.Cod_Cobert & " é obrigatório relacionar os bens e respectivos valores. A relação poderá ser preenchida diretamente no SYAS"
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
                lCountLine = lCountLine + 1
                lAuxiliar2 = "ou ser entregue juntamente com a proposta de seguro."
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
                lCountLine = lCountLine + 1
            End If
        Next lobjstcCob

        Do While lCountLine < 28
            Print #pNumArqDados, " "
            lCountLine = lCountLine + 1
        Loop

        GoTo ImprimeUsuario
    End If

    Do While lCountLine < 28
        Print #pNumArqDados, " "
        lCountLine = lCountLine + 1
    Loop

ImprimeUsuario:
    Print #pNumArqDados, pobjstcPed.Cod_User_Liber & " - " & gM46V999.gfForData(2, pobjstcPed.Dat_Liber) & " - " & Mid(Format(pobjstcPed.Hor_Liber, "0000"), 1, 2) & ":" & Mid(Format(pobjstcPed.Hor_Liber, "0000"), 3, 2)

    mfCriDados_MULTCOT1 = True
End Function

Private Function mfCriDados_MULTCOT2(ByVal pNumArqDados As Integer, ByVal pobjstcPed As stcA46V702B, _
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
    Dim lObjstcCorr As stcA46V705B    'Objeto stcA46V705B  (corretor).
    Dim lobjstcJur As stcA46V711B
    Dim lobjstcCob As stcA46V704B    'Objeto da classe (stc) de cobertura.
    Dim lCountLine As Integer
    Dim lValorParc As Variant
    Dim lValorComis As Single
    Dim lMensagem As String
    Dim lDescImportante() As String

    mfCriDados_MULTCOT2 = False

    If pobjstcPed.Cod_Dizer = 0 Then
        '001. Título do documento.
        Print #pNumArqDados, "COTAÇÃO DE SEGURO"
    Else
        Print #pNumArqDados, "COTAÇÃO DE ENDOSSO"
    End If

    '002. Título do seguro (de acordo com a base de dados).
    '     Neste caso deve ser colocado YASUDA RESIDENCIAL, YASUDA CONDOMÍNIO ou YASUDA EMPRESARIAL.
    Print #pNumArqDados, "YASUDA " & UCase(pobjstcPed.gDesc_Ramo)

    '003. Número da página.
    Print #pNumArqDados, "Página 1/" & pobjstcPed.ColecaoPedLoc.Count + 1

    '004. Emissão.
    Print #pNumArqDados, "Emissão: " & Format$(Now, "dd/mm/yyyy hh:nn:ss")

    'Ficha 901223
    '001B. Código do Processo Susep.
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

    '005. Validade.
    If pobjstcPedLoc.Tip_Emissao = e_TipEmissao_RenYasuda Then
        If gM46V999.gCodProduto >= 978 Then
            Print #pNumArqDados, "Validade: " & Format$(CDate(gM46V999.gfForData(2, pobjstcPed.Dat_Pri_Calculo)) + 15, "dd/mm/yyyy")
        Else
            Print #pNumArqDados, "Validade: " & Format$(CDate(gM46V999.gfForData(2, pobjstcPed.Dat_Pri_Calculo)) + 30, "dd/mm/yyyy")
        End If
    Else
        Print #pNumArqDados, "Validade: " & Format$(CDate(gM46V999.gfForData(2, pobjstcPed.Dat_Pri_Calculo)) + 15, "dd/mm/yyyy")
    End If

    '006. Número do pedido .
    lAuxiliar1 = pobjstcPed.Num_Ped
    Print #pNumArqDados, Format$(lAuxiliar1, "000000")

    '007. Corretor (código e nome) .
    If pobjstcPed.ColecaoPedCorr.Count > 0 Then
        Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(1)
        lAuxiliar1 = lObjstcCorr.Cod_Corr
        lAuxiliar2 = lObjstcCorr.gNom_Corretor
    Else
        lAuxiliar1 = 0
        lAuxiliar2 = " "
    End If

    'Soma da Comissão para saber se o prêmio é NET
    If pobjstcPed.ColecaoPedCorr.Count > 0 Then
        Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(1)
        lValorComis = lObjstcCorr.Per_Corr
        If pobjstcPed.ColecaoPedCorr.Count > 1 Then
            Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(2)
            lValorComis = lValorComis + lObjstcCorr.Per_Corr
            If pobjstcPed.ColecaoPedCorr.Count > 2 Then
                Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(3)
                lValorComis = lValorComis + lObjstcCorr.Per_Corr
            End If
        End If
    End If


    Print #pNumArqDados, Format$(lAuxiliar1, "0000") & " - " & lAuxiliar2

    '008. Tarifa .
    lAuxiliar2 = pobjstcPed.Cod_Produto
    If pobjstcPed.Cod_Produto = 976 And pobjstcPed.Cod_Ramo = e_CodRamo_Empresarial Then
        lAuxiliar2 = lAuxiliar2 & " A"
    End If
    If pobjstcPed.Cod_Sequencia = 1 Then
        lAuxiliar2 = lAuxiliar2 & " (A)"
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '009. Nosso número .
    lAuxiliar2 = pobjstcPed.NOSSO_NUMERO
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    If pobjstcPed.objstcCtrlEmis.Tip_Cotac = 2 And pobjstcPed.objstcCtrlEmis.Num_Cotac <> 0 Then
        Print #pNumArqDados, "Siscota: " & pobjstcPed.objstcCtrlEmis.Num_Cotac
    Else
        Print #pNumArqDados, " "
    End If

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
        'lAuxiliar2 = "CPF...........:"
        lAuxiliar2 = "               "
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        lAuxiliar2 = pobjstcPed.gCNPJ_CPF_Formatado & "   " & _
                     pobjstcPed.gCod_Doc_Segur & ": " & pobjstcPed.Num_Doc & "   " & _
                     "Órgão emissor: " & pobjstcPed.Org_Exp & "   " & _
                     "Emissão: " & gM46V999.gfForData(2, pobjstcPed.Dat_Emis_Doc)
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    Case e_TipPessoa_Juridica
        Print #pNumArqDados, "Jurídica"
        '013 e 014. Proponente - CNPJ ou CPF .
        'lAuxiliar2 = "CNPJ.........:"
        lAuxiliar2 = "              "
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
    'Ficha 900417
    Do While Len(lAuxiliar2) < 70
        lAuxiliar2 = lAuxiliar2 & " "
    Loop

    If gM46V999.gCodProduto >= 968 Then
        lAuxiliar2 = lAuxiliar2 & "Profissão: " & pobjstcPed.Nom_Empresa
    Else
        lAuxiliar2 = lAuxiliar2 & "Empresa: " & pobjstcPed.Nom_Empresa
    End If
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
    If lValorComis = 0 Then
        Print #pNumArqDados, "          Prêmio NET:"
    Else
        Print #pNumArqDados, "Prêmio Líquido Total:"
    End If
    lAuxiliar1 = pobjstcPed.VAL_PR
    lAuxiliar2 = Format$(lAuxiliar1, "standard")
    lTamanho = Len(lAuxiliar2)
    If lTamanho < 14 Then
        lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '*******************************************************************************
    If lValorComis = 0 Then
        'Caso seja prêmio net só irá imprimir as observações
        For lI = 1 To 30
            Print #pNumArqDados, " "
        Next
        GoTo ImprimirObservacao
    End If

    '********************************************************************************
    Print #pNumArqDados, "    Custo da Apólice:"
    '023. Resumo: custo da apólice.
    lAuxiliar1 = pobjstcPed.Val_Cust_Emis
    lAuxiliar2 = Format$(lAuxiliar1, "standard")
    lTamanho = Len(lAuxiliar2)
    If lTamanho < 14 Then
        lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    lAuxiliar2 = "IOF (" & Format(gM46V111.gProJurIOF, "standard") & "%):"
    Print #pNumArqDados, String(21 - Len(lAuxiliar2), " ") & lAuxiliar2
    '024. Resumo: IOF.
    lAuxiliar1 = pobjstcPed.Val_IOF
    lAuxiliar2 = Format$(lAuxiliar1, "standard")
    lTamanho = Len(lAuxiliar2)
    If lTamanho < 14 Then
        lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)


    Print #pNumArqDados, "Prêmio total a vista:"
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
    For lI = 1 To 12
        If lI > pobjstcPed.Qtd_Parc Then
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
                Set lobjstcJur = pobjstcPed.ColecaoPedJuros.Item(pobjstcPed.NOSSO_NUMERO & Format(lI, "00"))
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
    For lI = 1 To 12
        If lI > pobjstcPed.Qtd_Parc Then
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
                Set lobjstcJur = pobjstcPed.ColecaoPedJuros.Item(pobjstcPed.NOSSO_NUMERO & Format(lI, "00"))
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

    'Caso seja prêmio net virá direto para cá
    '********************************************************************************

ImprimirObservacao:

    '050 a 059. Importante.
    lCountLine = 0
    Call mpTextoImportante(pobjstcPedLoc, pobjstcPed, lDescImportante, "MULTCOT2", 36)
    For lI = 0 To 35
        If lDescImportante(lI) = "" Then Exit For
        lAuxiliar2 = lDescImportante(lI)
        If gM46V999.gfPreenchido(lAuxiliar2) Then
            lCountLine = lCountLine + 1
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        End If
    Next lI
    Call mpTextoObs(pobjstcPedLoc, lDescImportante, 36 - lCountLine)
    For lI = 0 To UBound(lDescImportante) - 1
        If lDescImportante(lI) = "" Then Exit For
        lAuxiliar2 = lDescImportante(lI)
        If gM46V999.gfPreenchido(lAuxiliar2) Then
            lCountLine = lCountLine + 1
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        End If
    Next lI

    If gM46V999.gCodProduto >= 958 And pobjstcPedLoc.Cob_Bens = True And Not mfVerDigBens() Then
        For Each lobjstcCob In pobjstcPedLoc.ColecaoPedCobert
            If gM46V999.gCod_Ramo = e_CodRamo_Empresarial And lobjstcCob.Cod_Cobert = gM46V999.gfBuscarDePara(94, 1) Then
                lAuxiliar2 = "Para a cobertura " & lobjstcCob.Cod_Cobert & " é obrigatório relacionar os bens e respectivos valores. A relação poderá ser preenchida diretamente no SYAS"
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
                lCountLine = lCountLine + 1
                lAuxiliar2 = "ou ser entregue juntamente com a proposta de seguro."
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
                lCountLine = lCountLine + 1
            End If
            If gM46V999.gCod_Ramo = e_CodRamo_Residencial And lobjstcCob.Cod_Cobert = gM46V999.gfBuscarDePara(27, 1) Or _
               lobjstcCob.Cod_Cobert = gM46V999.gfBuscarDePara(29, 1) Or _
               lobjstcCob.Cod_Cobert = gM46V999.gfBuscarDePara(30, 1) Or _
               lobjstcCob.Cod_Cobert = gM46V999.gfBuscarDePara(31, 1) Then
                lAuxiliar2 = "Para a cobertura " & lobjstcCob.Cod_Cobert & " é obrigatório relacionar os bens e respectivos valores. A relação poderá ser preenchida diretamente no SYAS"
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
                lCountLine = lCountLine + 1
                lAuxiliar2 = "ou ser entregue juntamente com a proposta de seguro."
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
                lCountLine = lCountLine + 1
            End If
        Next lobjstcCob
        Do While lCountLine < 36
            Print #pNumArqDados, " "
            lCountLine = lCountLine + 1
        Loop
        GoTo ImprimeUsuario
    End If

    Do While lCountLine < 36
        Print #pNumArqDados, " "
        lCountLine = lCountLine + 1
    Loop

ImprimeUsuario:
    Print #pNumArqDados, pobjstcPed.Cod_User_Liber & " - " & gM46V999.gfForData(2, pobjstcPed.Dat_Liber) & " - " & Mid(Format(pobjstcPed.Hor_Liber, "0000"), 1, 2) & ":" & Mid(Format(pobjstcPed.Hor_Liber, "0000"), 3, 2)
    mfCriDados_MULTCOT2 = True
End Function


Private Function mfCriDados_MULTCOT3(ByVal pNumArqDados As Integer, ByVal pobjstcPed As stcA46V702B, _
                                     pNumDe As Integer, _
                                     pNumAte As Integer) As Boolean
'Função: cria arquivo de dados - cotação coletiva - dados do item -

'Parâmetro de entrada...pNumArqDados...Número do arquivo de dados.

    Dim lAuxiliar1 As Double  'Auxiliar - utilizado apenas como exemplo.
    Dim lAuxiliar2 As String  'Auxiliar - utilizado apenas como exemplo.
    Dim lAuxiliar3 As String  'Auxiliar - utilizado apenas como exemplo.
    Dim lCodCobertura As Integer    'Código da cobertura.
    Dim lI As Integer    'Utilizado com For...Next.
    Dim lTamanho As Integer    'Tamanho.
    Dim lObjstcCorr As stcA46V705B
    Dim lobjstcCob As stcA46V704B    'Objeto da classe (stc) de cobertura.
    Dim lobjstcSeg As stcA46V716B    'Outra seguradora.
    Dim lobjstcBem As stcA46V715B    'Bens
    Dim lCountLine As Integer
    Dim lValTotCober As Double    'Soma das coberturas.
    Dim lDescFranquias() As String
    Dim lobjstcPedLoc As stcA46V708B
    Dim lCount As Integer
    Dim lPrLiqItem As Double
    Dim lPrazo As Integer
    Dim lCodCobert As Integer
    Dim lNumMeses As Integer
    Dim lDescCom As Double
    Dim lDescExp As Double
    Dim lbdTab_Tipo_Fator As ADODB.Connection   'Arquivo: P00Tarifa.
    Dim lrsTab_Tipo_Fator As clsYasRecordSet    'Registro :Tab_Tipo_Fator.
    Dim lbdTbm_Regiao As ADODB.Connection
    Dim lrsTbm_Regiao As clsYasRecordSet
    Dim lobjstcFator As stcA46V720B       'Objeto da classe de estrutura stcA46V720B (TAB_PED_FAT_RISCO).
    Dim lblnAchou As Boolean        'Indica se achou  o fator ou não.
    Dim lMensagem As String
    Dim lDemParcelas As String    'Demais parcelas.
    Dim lQuaDias As Integer
    Dim lSelect As String


    mfCriDados_MULTCOT3 = False
    For Each lobjstcPedLoc In pobjstcPed.ColecaoPedLoc
        '001. Título do documento.
        If lobjstcPedLoc.Cod_DizerItem = 0 Then
            Print #pNumArqDados, "ANEXO DA COTAÇÃO DE SEGURO"
        Else
            Print #pNumArqDados, "ANEXO DA COTAÇÃO DE ENDOSSO"
        End If

        '002. Título do seguro (de acordo com a base de dados).
        '     Neste caso deve ser colocado YASUDA RESIDENCIAL, YASUDA CONDOMÍNIO ou YASUDA EMPRESARIAL.
        Print #pNumArqDados, "YASUDA " & UCase(pobjstcPed.gDesc_Ramo)

        '003. Número da página.
        Print #pNumArqDados, "Página "; pNumDe & "/" & pNumAte
        pNumDe = pNumDe + 1

        '004. Emissão.
        Print #pNumArqDados, "Emissão: " & Format$(Now, "dd/mm/yyyy hh:nn:ss")

        'Ficha 901223
        '001B. Código do Processo Susep.
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

        '005. Validade.
        If lobjstcPedLoc.Tip_Emissao = e_TipEmissao_RenYasuda Then
            If gM46V999.gCodProduto >= 978 Then
                Print #pNumArqDados, "Validade: " & Format$(CDate(gM46V999.gfForData(2, pobjstcPed.Dat_Pri_Calculo)) + 15, "dd/mm/yyyy")
            Else
                Print #pNumArqDados, "Validade: " & Format$(CDate(gM46V999.gfForData(2, pobjstcPed.Dat_Pri_Calculo)) + 30, "dd/mm/yyyy")
            End If
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
        If pobjstcPed.ColecaoPedCorr.Count > 0 Then
            Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(1)
            lAuxiliar1 = lObjstcCorr.Cod_Corr
            lAuxiliar2 = lObjstcCorr.gNom_Corretor
        Else
            lAuxiliar1 = 0
            lAuxiliar2 = " "
        End If
        Print #pNumArqDados, Format$(lAuxiliar1, "0000") & " - " & lAuxiliar2

        '009. Tarifa .
        lAuxiliar2 = pobjstcPed.Cod_Produto
        If pobjstcPed.Cod_Produto = 976 And pobjstcPed.Cod_Ramo = e_CodRamo_Empresarial Then
            lAuxiliar2 = lAuxiliar2 & " A"
        End If
        If pobjstcPed.Cod_Sequencia = 1 Then
            lAuxiliar2 = lAuxiliar2 & " (A)"
        End If
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
        If gM46V999.gCod_Ramo = e_CodRamo_Empresarial Then
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
        'lAuxiliar2 = lObjstcPedLoc.Cod_Regiao & " - " & Left(lObjstcPedLoc.gDesc_Regiao, 38)
        'Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        If pobjstcPed.Cod_Produto <= gM46V999.TarifaAntiga Then
            lAuxiliar2 = lobjstcPedLoc.Cod_Regiao & " - " & Left(lobjstcPedLoc.gDesc_Regiao, 38)
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        Else
            'Nova Região
            If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTbm_Regiao, lMensagem) = False Then
                Call gM46V999.gpGraLog(2, lMensagem)
                Exit Function
            End If

            lSelect = "Select Cod_Faixa_Reg, DSC_REG from Tbm_Faixa_Regiao "
            lSelect = lSelect & " Where cod_Ramo = " & pobjstcPed.Cod_Ramo
            lSelect = lSelect & "    and Cod_Modalidade = " & lobjstcPedLoc.Cod_Produto
            lSelect = lSelect & "    and Cod_Tarif = " & pobjstcPed.Cod_Produto
            lSelect = lSelect & "    and COD_Sequencia <= " & pobjstcPed.Cod_Sequencia
            lSelect = lSelect & "    and Cep_Ini <= " & lobjstcPedLoc.Num_Cep_Risco
            lSelect = lSelect & "    and Cep_Fim >= " & lobjstcPedLoc.Num_Cep_Risco
            lSelect = lSelect & "    order by Cod_Sequencia desc "

            If gM46V999.gfObtRegistro(lbdTbm_Regiao, lSelect, lrsTbm_Regiao, lMensagem) = False Then
                Call gM46V999.gpGraLog(0, lMensagem)
                Call gM46V999.gpFecha2(lbdTbm_Regiao)
                Exit Function
            End If

            If Not lrsTbm_Regiao.EOF Then
                lAuxiliar2 = gM46V999.Check(lrsTbm_Regiao("Cod_Faixa_Reg")) & " - " & Left(gM46V999.Check(lrsTbm_Regiao("DSC_REG")), 38)
            Else
                lAuxiliar2 = "Não cadastrada."
            End If
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        End If

        '023. Dados do seguro - vigência .
        lAuxiliar2 = "de " & gM46V999.gfForData(2, pobjstcPed.Dat_Ini_Vig) & _
                   " a " & gM46V999.gfForData(2, pobjstcPed.Dat_Ter_Vig)

        'Se for perfil Calculo não será impresso conforme ficha 0701831
        If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda And _
           (gM46V999.PerfilAcesso = e_Perfil_CALC Or gM46V999.PerfilAcesso = e_Perfil_CALC_Gestor) Then
            lQuaDias = gM46V999.gfDifDatas(pobjstcPed.Dat_Ini_Vig, pobjstcPed.Dat_Ter_Vig)
            If lQuaDias >= 365 Then
                lAuxiliar2 = ""
            End If
        End If
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '024. Dados do seguro - valor em risco .
        lAuxiliar2 = "R$ " & Format(lobjstcPedLoc.Val_Vr, "standard")
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '025. Dados do seguro - tipo de emissão .
        lAuxiliar2 = lobjstcPedLoc.gDesc_TipEmissao
        Select Case lobjstcPedLoc.Tip_Emissao
        Case e_TipEmissao_RenYasuda, e_TipEmissao_Endosso
            lAuxiliar2 = lAuxiliar2 & "  Apólice: " & lobjstcPedLoc.Num_Apol_Anterior & "  Item: " & lobjstcPedLoc.Num_Item_Anterior & "  Sinistro: " & IIf(lobjstcPedLoc.Ind_Sinistro = e_SimNao_Sim, "Sim", "Não")
        End Select
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '026 e 027. Coberturas contratadas - título.
        'm       If pObjstcPed.Cod_Ramo = e_CodRamo_Empresarial Then
        'm            '035 e 036. Coberturas contratadas - título.
        'm            lAuxiliar2 = "Desconto de"
        'm            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'm            lAuxiliar2 = "experiência"
        'm            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'm        Else
        Print #pNumArqDados, " "
        Print #pNumArqDados, " "
        'm        End If


        '028 a 202. Coberturas contratadas - linhas de 01 a 22.
        lCountLine = 0
        lValTotCober = 0
        For Each lobjstcCob In lobjstcPedLoc.ColecaoPedCobert
            lCountLine = lCountLine + 1
            'Código
            Print #pNumArqDados, Format$(lobjstcCob.Cod_Cobert, "0000")
            lCodCobert = lobjstcCob.Cod_Cobert
            'Descrição
            Print #pNumArqDados, Left(lobjstcCob.gNomeCobertura, 45)

            lAuxiliar1 = IIf(IsNumeric(lobjstcCob.Num_Meses_PI), lobjstcCob.Num_Meses_PI, 0)
            If lAuxiliar1 = 0 Then
                Print #pNumArqDados, " "
                Print #pNumArqDados, " "
            Else
                'pi mes
                Print #pNumArqDados, Format$(lobjstcCob.Num_Meses_PI, "00")
                lNumMeses = lobjstcCob.Num_Meses_PI
                Call mpPrazoTabCoefPI(pobjstcPed, lCodCobert, lNumMeses, lPrazo)
                'pi dias
                If lPrazo <> 0 Then
                    Print #pNumArqDados, lPrazo
                Else
                    Print #pNumArqDados, " "
                End If
            End If
            '%Sobre Prejuízos
            lAuxiliar1 = lobjstcCob.Per_Franq
            lAuxiliar2 = Format$(lAuxiliar1, "standard")
            lTamanho = Len(lAuxiliar2)
            If lTamanho < 10 Then
                lAuxiliar2 = Space(10 - lTamanho) & lAuxiliar2
            End If
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

            'Limitado ao Mínimo
            lAuxiliar1 = lobjstcCob.Val_Franq_Min
            lAuxiliar2 = Format$(lAuxiliar1, "standard")
            lTamanho = Len(lAuxiliar2)
            If lTamanho < 14 Then
                lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
            End If
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

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
            lDescCom = lobjstcCob.Per_Desc_Coml
            lAuxiliar2 = Format$(lAuxiliar1, "standard")
            lTamanho = Len(lAuxiliar2)
            If lTamanho < 10 Then
                lAuxiliar2 = Space(10 - lTamanho) & lAuxiliar2
            End If

            If pobjstcPed.Cod_Ramo = e_CodRamo_Empresarial Then
                'e) Desconto de experiência .
                lAuxiliar1 = lobjstcCob.Per_Desc_Exp
                lDescExp = lobjstcCob.Per_Desc_Exp
                lAuxiliar2 = Format$(lAuxiliar1, "standard")
                lTamanho = Len(lAuxiliar2)
                If lTamanho < 10 Then
                    lAuxiliar2 = Space(10 - lTamanho) & lAuxiliar2
                End If
            Else
                lDescExp = 0
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
            Print #pNumArqDados, " "
            Print #pNumArqDados, " "
        Loop

        '203. Coberturas contratadas - soma.
        lAuxiliar1 = lValTotCober
        lAuxiliar2 = Format$(lAuxiliar1, "standard")
        lTamanho = Len(lAuxiliar2)
        If lTamanho < 14 Then
            lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
        End If
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '204. Coberturas contratadas - prêmio líquido.
        If lValTotCober < gM46V999.gProPremioMinimo Then
            lValTotCober = gM46V999.gProPremioMinimo
        End If

        lAuxiliar1 = lValTotCober
        lAuxiliar2 = Format$(lAuxiliar1, "standard")
        lTamanho = Len(lAuxiliar2)
        If lTamanho < 14 Then
            lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
        End If
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '205. Informação cifrada .
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

        '*   lAuxiliar2 = lAuxiliar2 & "/A:" & Format(lObjstcPedLoc.Per_Desc_Apar, "00")

        'Ficha 900740 e 901673 e 901772
        If (gM46V111.gCod_plano >= 2021 And gM46V111.gCod_plano <= 2031) Or (gM46V111.gCod_plano >= 2061 And gM46V111.gCod_plano <= 2100) Or (gM46V111.gCod_plano >= 2111 And gM46V111.gCod_plano <= 2130) Then
            lAuxiliar2 = lAuxiliar2 & "/P:00"
        Else
            lAuxiliar2 = lAuxiliar2 & "/P:" & Format(lobjstcPedLoc.Per_Desc_Progres, "00")
        End If

        'lAuxiliar2 = lAuxiliar2 & "/CO:" & Format(pObjstcPed.objstcCtrlEmis.Per_Corr + pObjstcPed.objstcCtrlEmis.Per_Corr_2 + pObjstcPed.objstcCtrlEmis.Per_Corr_3, "00.00")
        lAuxiliar1 = 0
        If pobjstcPed.ColecaoPedCorr.Count > 0 Then
            Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(1)
            lAuxiliar1 = lObjstcCorr.Per_Corr
            If pobjstcPed.ColecaoPedCorr.Count > 1 Then
                Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(2)
                lAuxiliar1 = lAuxiliar1 + lObjstcCorr.Per_Corr
                If pobjstcPed.ColecaoPedCorr.Count > 2 Then
                    Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(3)
                    lAuxiliar1 = lAuxiliar1 + lObjstcCorr.Per_Corr
                End If
            End If
        End If
        lAuxiliar2 = lAuxiliar2 & "/CO:" & Format(lAuxiliar1 * 100, "0000")
        '*   lAuxiliar2 = lAuxiliar2 & "/C:" & Format(Round(lObjstcPedLoc.Val_Coef_IS_Vr, 4) * 10000, "00000")

        'Ficha 900740 e 901673 e 901772
        If (gM46V111.gCod_plano >= 2021 And gM46V111.gCod_plano <= 2031) Or (gM46V111.gCod_plano >= 2061 And gM46V111.gCod_plano <= 2100) Or (gM46V111.gCod_plano >= 2111 And gM46V111.gCod_plano <= 2130) Then
            lAuxiliar2 = lAuxiliar2 & "/DC:00"
        Else
            lAuxiliar2 = lAuxiliar2 & "/DC:" & Format(lDescCom * 100, "0000")
        End If

        '*   lAuxiliar2 = lAuxiliar2 & "/DE:" & Format(lDescExp, "00")
        '   lAuxiliar2 = lAuxiliar2 & "/PRGRS6"

        'Ficha 900550
        '/V1.0 -> versão após ajustar o bônus das coberturas

        lAuxiliar2 = lAuxiliar2 & "/V1.0"
        lAuxiliar2 = lAuxiliar2 & "/" & Format(gM46V111.gProJurIOF * 100)


        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '206 a 213. Franquias.
        Print #pNumArqDados, "FRANQUIAS"
        lCountLine = 1
        'Call mpDescricaoFranquia_Corretor(lObjstcPedLoc, pobjstcPed, lDescFranquias, 51)

        ''Não deve mais buscar pois já está gravado
        'Call gM46V999.mpNovaBuscaDescricaoFranquia(pobjstcPed, lObjstcPedLoc, 51, 0, lDescFranquias)

        If UBound(lDescFranquias) > 0 Then
            For lI = 0 To 50
                If lDescFranquias(lI) = "" Then Exit For
                lAuxiliar2 = lDescFranquias(lI)
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
                lCountLine = lCountLine + 1
            Next lI
        End If

        'Call mpDescricaoFranquia(lobjstcPedLoc, pobjstcPed, lDescFranquias, 51 - lCountLine)
        'If UBound(lDescFranquias) > 0 Then
        '    For li = 0 To 50 - lCountLine
        '        If lDescFranquias(li) = "" Then Exit For
        '        lAuxiliar2 = lDescFranquias(li)
        '        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        '        lCountLine = lCountLine + 1
        '    Next li
        'End If

        '        '214 a 238. Fatores de Proteção - linhas de 01 a 25.
        '        lCountLine = 0
        '        If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTab_Tipo_Fator, lMensagem) = False Then
        '            Call gM46V999.gpGraLog(2, lMensagem)
        '            Exit Function
        '        End If
        '
        '        lSelect = " Select distinct DSC_FATOR, COD_FATOR  from Tbm_Tipo_Fator_Protec "
        '        lSelect = lSelect & " Where cod_Ramo = " & pobjstcPed.Cod_Ramo
        '        lSelect = lSelect & " and Cod_Modalidade = " & lobjstcPedLoc.Cod_Produto
        '        lSelect = lSelect & " and COD_TARIF = " & pobjstcPed.Cod_Produto
        '        lSelect = lSelect & " and COD_Sequencia <= " & pobjstcPed.Cod_Sequencia
        '        lSelect = lSelect & " order by DSC_FATOR"
        '
        '        If gM46V999.gfObtRegistro(lbdTab_Tipo_Fator, lSelect, lrsTab_Tipo_Fator, lMensagem) = False Then
        '            Call gM46V999.gpFecha2(lbdTab_Tipo_Fator)
        '            Call gM46V999.gpGraLog(2, lMensagem)
        '            Screen.MousePointer = vbDefault
        '            Exit Function
        '        End If
        '
        '        Do While Not lrsTab_Tipo_Fator.EOF
        '            lCountLine = lCountLine + 1
        '            lblnAchou = False
        '
        '            For Each lobjstcFator In lobjstcPedLoc.ColecaoPedFator
        '                If lobjstcFator.Cod_Fator = lrsTab_Tipo_Fator("Cod_Fator") Then
        '                    lblnAchou = True
        '                    lAuxiliar2 = Format(lCountLine, "00") & " (X) " & Trim(lrsTab_Tipo_Fator("Dsc_Fator"))
        '                    If Trim(lobjstcFator.Dsc_Compl_Fator) <> "" Then
        '                        lAuxiliar2 = lAuxiliar2 & " - " & Trim(lobjstcFator.Dsc_Compl_Fator)
        '                    End If
        '                    Print #pNumArqDados, lAuxiliar2
        '                    Exit For
        '                End If
        '            Next lobjstcFator
        '            If lblnAchou = False Then
        '                lAuxiliar2 = Format(lCountLine, "00") & " ( ) " & Trim(lrsTab_Tipo_Fator("Dsc_Fator"))
        '                Print #pNumArqDados, lAuxiliar2
        '            End If
        '            lrsTab_Tipo_Fator.MoveNext
        '        Loop
        '
        '        Call gM46V999.gpFecha2(lbdTab_Tipo_Fator)
        '
        '        Do While lCountLine < 14
        '            lCountLine = lCountLine + 1
        '            Print #pNumArqDados, " "
        '        Loop
        '
        '
        '        lCountLine = 0
        '        For Each lobjstcBem In lobjstcPedLoc.ColecaoPedLocBem
        '            lCountLine = lCountLine + 1
        '            lAuxiliar3 = lobjstcBem.gTipBem & " - " & lobjstcBem.Dsc_Bem
        '            lAuxiliar2 = Format(lCountLine, "000") & " - " & Left(lAuxiliar3, 100) & String(90 - Len(Left(lAuxiliar3, 90)), ".") & "R$ " & Format(lobjstcBem.Val_Bem, "standard")
        '            Print #pNumArqDados, lAuxiliar2
        '        Next lobjstcBem

        Do While lCountLine < 52
            lCountLine = lCountLine + 1
            Print #pNumArqDados, " "
        Loop
    Next lobjstcPedLoc


    mfCriDados_MULTCOT3 = True

End Function


Private Function mfCriDados_MULTPRO1(ByVal pNumArqDados As Integer, ByVal pobjstcPed As stcA46V702B, _
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
    Dim lPrazo As Integer
    Dim lCodCobert As Integer
    Dim lNumMeses As Integer
    Dim lDescCom As Double
    Dim lDescExp As Double
    Dim lValorParc As Variant
    Dim lMensagem As String
    Dim lOrdem As String    'Ordem.
    Dim lPriParcela As String   'Primeira parcela.
    Dim lDemParcelas As String    'Demais parcelas.
    Dim lJuros As String    'Juros.
    Dim lbdTbm_Regiao As ADODB.Connection
    Dim lrsTbm_Regiao As clsYasRecordSet
    Dim lSelect As String

    mfCriDados_MULTPRO1 = False

    '001. Título do seguro (de acordo com a base de dados).
    '     Neste caso deve ser colocado PROPOSTA DE SEGURO YASUDA...
    '     RESIDENCIAL, CONDOMÍNIO ou EMPRESARIAL.
    '**********
    'Print #pNumArqDados, "PROPOSTA DE SEGURO YASUDA " & UCase(pObjstcPed.gDesc_Ramo)
    '**********
    Print #pNumArqDados, UCase(pobjstcPed.gDesc_Ramo)


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
    '**************
    'Print #pNumArqDados, "Emissão: " & Format$(Now, "dd/mm/yyyy hh:nn:ss")
    '**************
    Print #pNumArqDados, Format$(Now, "dd/mm/yyyy hh:nn:ss")
    '004. Número do pedido .
    lAuxiliar1 = pobjstcPed.Num_Ped
    Print #pNumArqDados, Format$(lAuxiliar1, "000000")

    '005. Número do item .
    lAuxiliar1 = pobjstcPedLoc.Num_Item
    Print #pNumArqDados, Format$(lAuxiliar1, "0000")

    '006. Corretor (código e nome) .
    If pobjstcPed.ColecaoPedCorr.Count > 0 Then
        Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(1)
        lAuxiliar1 = lObjstcCorr.Cod_Corr
        lAuxiliar2 = lObjstcCorr.gNom_Corretor
    Else
        lAuxiliar1 = 0
        lAuxiliar2 = " "
    End If
    Print #pNumArqDados, Format$(lAuxiliar1, "0000") & " - " & lAuxiliar2

    '007. Tarifa .
    lAuxiliar2 = pobjstcPed.Cod_Produto
    If pobjstcPed.Cod_Produto = 976 And pobjstcPed.Cod_Ramo = e_CodRamo_Empresarial Then
        lAuxiliar2 = lAuxiliar2 & " A"
    End If
    If pobjstcPed.Cod_Sequencia = 1 Then
        lAuxiliar2 = lAuxiliar2 & " (A)"
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '008. Nosso número .
    lAuxiliar2 = pobjstcPed.NOSSO_NUMERO
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    If pobjstcPed.objstcCtrlEmis.Tip_Cotac = 2 And pobjstcPed.objstcCtrlEmis.Num_Cotac <> 0 Then
        Print #pNumArqDados, "Siscota: " & pobjstcPed.objstcCtrlEmis.Num_Cotac
    Else
        Print #pNumArqDados, " "
    End If

    Print #pNumArqDados, pobjstcPed.Cod_Unicidade & " - " & pobjstcPed.Cod_Ver_Unici

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
        'lAuxiliar2 = "CNPJ.........:"
        lAuxiliar2 = "              "
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        lAuxiliar2 = pobjstcPed.gCNPJ_CPF_Formatado & "   " & _
                     "Ramo de atividade: " & pobjstcPed.gDesc_RamoAtiv
        If pobjstcPed.Val_RendaFatu > 0 Then
            lAuxiliar2 = lAuxiliar2 & " Faturamento: " & Format(pobjstcPed.Val_RendaFatu, "standard")
        Else
            lAuxiliar2 = lAuxiliar2 & " Faturamento: não informado"
        End If
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    Case e_TipPessoa_Fisica
        Print #pNumArqDados, "Física"
        '012 e 013. Proponente - CNPJ ou CPF .
        'lAuxiliar2 = "CPF...........:"
        lAuxiliar2 = "               "
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        lAuxiliar2 = pobjstcPed.gCNPJ_CPF_Formatado & "   " & _
                     pobjstcPed.gCod_Doc_Segur & ": " & Trim(pobjstcPed.Num_Doc) & "   " & _
                     "Órgão emissor: " & Trim(pobjstcPed.Org_Exp) & "   " & _
                     "Emissão: " & gM46V999.gfForData(2, pobjstcPed.Dat_Emis_Doc)

        If pobjstcPed.Val_RendaFatu > 0 Then
            lAuxiliar2 = lAuxiliar2 & " Renda: " & Format(pobjstcPed.Val_RendaFatu, "standard")
        Else
            lAuxiliar2 = lAuxiliar2 & " Renda: não informada "
        End If

        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    End Select

    '014. Proponente - endereço .
    lAuxiliar2 = pobjstcPed.objstcPedEndereco.Cod_Tip_Logr & " " & _
                 pobjstcPed.objstcPedEndereco.Nom_Logr & "," & _
                 pobjstcPed.objstcPedEndereco.Dsc_Compl_Logr
    'Ficha 900417
    Do While Len(lAuxiliar2) < 70
        lAuxiliar2 = lAuxiliar2 & " "
    Loop

    If gM46V999.gCodProduto >= 968 Then
        lAuxiliar2 = lAuxiliar2 & "Profissão: " & pobjstcPed.Nom_Empresa
    Else
        lAuxiliar2 = lAuxiliar2 & "Empresa: " & pobjstcPed.Nom_Empresa
    End If
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
    If gM46V999.gCod_Ramo = e_CodRamo_Empresarial Then
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
    'lAuxiliar2 = pobjstcPedLoc.Cod_Regiao & " - " & Left(pobjstcPedLoc.gDesc_Regiao, 38)
    'Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    If pobjstcPed.Cod_Produto <= gM46V999.TarifaAntiga Then
        lAuxiliar2 = pobjstcPedLoc.Cod_Regiao & " - " & Left(pobjstcPedLoc.gDesc_Regiao, 38)
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    Else
        'Nova Região
        If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTbm_Regiao, lMensagem) = False Then
            Call gM46V999.gpGraLog(2, lMensagem)
            Exit Function
        End If

        lSelect = "Select Cod_Faixa_Reg, DSC_REG from Tbm_Faixa_Regiao "
        lSelect = lSelect & " Where cod_Ramo = " & pobjstcPed.Cod_Ramo
        lSelect = lSelect & "    and Cod_Modalidade = " & pobjstcPedLoc.Cod_Produto
        lSelect = lSelect & "    and Cod_Tarif = " & pobjstcPed.Cod_Produto
        lSelect = lSelect & "    and COD_Sequencia <= " & pobjstcPed.Cod_Sequencia
        lSelect = lSelect & "    and Cep_Ini <= " & pobjstcPedLoc.Num_Cep_Risco
        lSelect = lSelect & "    and Cep_Fim >= " & pobjstcPedLoc.Num_Cep_Risco
        lSelect = lSelect & "    order by Cod_Sequencia desc "

        If gM46V999.gfObtRegistro(lbdTbm_Regiao, lSelect, lrsTbm_Regiao, lMensagem) = False Then
            Call gM46V999.gpGraLog(0, lMensagem)
            Call gM46V999.gpFecha2(lbdTbm_Regiao)
            Exit Function
        End If

        If Not lrsTbm_Regiao.EOF Then
            lAuxiliar2 = gM46V999.Check(lrsTbm_Regiao("Cod_Faixa_Reg")) & " - " & Left(gM46V999.Check(lrsTbm_Regiao("DSC_REG")), 38)
        Else
            lAuxiliar2 = "Não cadastrada."
        End If
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    End If

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
    Case e_TipEmissao_RenYasuda, e_TipEmissao_Endosso
        lAuxiliar2 = lAuxiliar2 & "  Apólice: " & pobjstcPedLoc.Num_Apol_Anterior & "  Item: " & pobjstcPedLoc.Num_Item_Anterior & "  Sinistro: " & IIf(pobjstcPedLoc.Ind_Sinistro = e_SimNao_Sim, "Sim", "Não")
    End Select
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '035 e 036. Coberturas contratadas - desconto - título.
    'mIf pObjstcPed.Cod_Ramo = e_CodRamo_Empresarial Then
    'm    lAuxiliar2 = "Desconto de"
    'm    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    'm    lAuxiliar2 = "experiência"
    'm    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    'mElse
    Print #pNumArqDados, " "
    Print #pNumArqDados, " "
    'mEnd If

    '037 a 168. Coberturas contratadas - linhas de 01 a 22.
    lCountLine = 0
    lValTotCober = 0
    For Each lobjstcCob In pobjstcPedLoc.ColecaoPedCobert
        lCountLine = lCountLine + 1
        'Código
        Print #pNumArqDados, Format$(lobjstcCob.Cod_Cobert, "0000")
        lCodCobert = lobjstcCob.Cod_Cobert
        'Descrição
        Print #pNumArqDados, Left(lobjstcCob.gNomeCobertura, 45)

        lAuxiliar1 = IIf(IsNumeric(lobjstcCob.Num_Meses_PI), lobjstcCob.Num_Meses_PI, 0)
        If lAuxiliar1 = 0 Then
            Print #pNumArqDados, " "

            Print #pNumArqDados, " "
        Else
            'pi mes
            Print #pNumArqDados, Format$(lobjstcCob.Num_Meses_PI, "00")
            lNumMeses = lobjstcCob.Num_Meses_PI
            Call mpPrazoTabCoefPI(pobjstcPed, lCodCobert, lNumMeses, lPrazo)
            'pi dias
            If lPrazo <> 0 Then
                Print #pNumArqDados, lPrazo
            Else
                Print #pNumArqDados, " "
            End If
        End If
        '%Sobre Prejuízos
        lAuxiliar1 = lobjstcCob.Per_Franq
        lAuxiliar2 = Format$(lAuxiliar1, "standard")
        lTamanho = Len(lAuxiliar2)
        If lTamanho < 10 Then
            lAuxiliar2 = Space(10 - lTamanho) & lAuxiliar2
        End If
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        'Limitado ao Mínimo
        lAuxiliar1 = lobjstcCob.Val_Franq_Min
        lAuxiliar2 = Format$(lAuxiliar1, "standard")
        lTamanho = Len(lAuxiliar2)
        If lTamanho < 14 Then
            lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
        End If
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

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
        lDescCom = lobjstcCob.Per_Desc_Coml
        lAuxiliar2 = Format$(lAuxiliar1, "standard")
        lTamanho = Len(lAuxiliar2)
        If lTamanho < 10 Then
            lAuxiliar2 = Space(10 - lTamanho) & lAuxiliar2
        End If
        'm        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        If pobjstcPed.Cod_Ramo = e_CodRamo_Empresarial Then
            'e) Desconto de experiência .
            lAuxiliar1 = lobjstcCob.Per_Desc_Exp
            lAuxiliar2 = Format$(lAuxiliar1, "standard")
            lTamanho = Len(lAuxiliar2)
            If lTamanho < 10 Then
                lAuxiliar2 = Space(10 - lTamanho) & lAuxiliar2
            End If
            'm            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        Else
            lDescExp = 0
            'm            Print #pNumArqDados, " "
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

    '171 A. Coberturas contratadas - PER IOF.
    lAuxiliar2 = "IOF (" & Format(gM46V111.gProJurIOF, "standard") & "%):"
    lTamanho = Len(lAuxiliar2)
    lAuxiliar2 = Space(17 - lTamanho) & lAuxiliar2
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

    '* lAuxiliar2 = lAuxiliar2 & "/A:" & Format(pobjstcPedLoc.Per_Desc_Apar, "00")

    'Ficha 900740 e 901673 e 901772
    If (gM46V111.gCod_plano >= 2021 And gM46V111.gCod_plano <= 2031) Or (gM46V111.gCod_plano >= 2061 And gM46V111.gCod_plano <= 2100) Or (gM46V111.gCod_plano >= 2111 And gM46V111.gCod_plano <= 2130) Then
        lAuxiliar2 = lAuxiliar2 & "/P:00"
    Else
        lAuxiliar2 = lAuxiliar2 & "/P:" & Format(pobjstcPedLoc.Per_Desc_Progres, "00")
    End If

    'lAuxiliar2 = lAuxiliar2 & "/CO:" & Format(pObjstcPed.objstcCtrlEmis.Per_Corr + pObjstcPed.objstcCtrlEmis.Per_Corr_2 + pObjstcPed.objstcCtrlEmis.Per_Corr_3, "00.00")
    lAuxiliar1 = 0
    If pobjstcPed.ColecaoPedCorr.Count > 0 Then
        Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(1)
        lAuxiliar1 = lObjstcCorr.Per_Corr
        If pobjstcPed.ColecaoPedCorr.Count > 1 Then
            Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(2)
            lAuxiliar1 = lAuxiliar1 + lObjstcCorr.Per_Corr
            If pobjstcPed.ColecaoPedCorr.Count > 2 Then
                Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(3)
                lAuxiliar1 = lAuxiliar1 + lObjstcCorr.Per_Corr
            End If
        End If
    End If

    lAuxiliar2 = lAuxiliar2 & "/CO:" & Format(lAuxiliar1 * 100, "0000")
    '*  lAuxiliar2 = lAuxiliar2 & "/C:" & Format(Round(pobjstcPedLoc.Val_Coef_IS_Vr, 4) * 10000, "00000")

    'Ficha 900740 e 901673 e 901772
    If (gM46V111.gCod_plano >= 2021 And gM46V111.gCod_plano <= 2031) Or (gM46V111.gCod_plano >= 2061 And gM46V111.gCod_plano <= 2100) Or (gM46V111.gCod_plano >= 2111 And gM46V111.gCod_plano <= 2130) Then
        lAuxiliar2 = lAuxiliar2 & "/DC:00"
    Else
        lAuxiliar2 = lAuxiliar2 & "/DC:" & Format(lDescCom * 100, "0000")
    End If

    '*  lAuxiliar2 = lAuxiliar2 & "/DE:" & Format(lDescExp, "00")
    '  lAuxiliar2 = lAuxiliar2 & "/PRGRS6"

    'Ficha 900550
    '/V1.0 -> versão após ajustar o bônus das coberturas

    lAuxiliar2 = lAuxiliar2 & "/V1.0"

    '###INSPECAO
    If gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Or gM46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda Then
        If pobjstcPedLoc.Insp_Obr = True Then
            If pobjstcPedLoc.Insp_Agend = True Then
                lAuxiliar2 = lAuxiliar2 & "/INSPS"
            Else
                lAuxiliar2 = lAuxiliar2 & "/INSPN"
            End If
        End If
    End If

    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    'Calcula o fracionamento
    With pobjstcPed
        Call gM46V998.gfCalFracionamento(pobjstcPed, pobjstcPedLoc.Tip_Emissao, _
                                         lValorParc, lMensagem, gM46V111.gProJurIOF, .Cod_Reg, .Val_IOF, .Val_Parc_Min)
    End With

    '219 a 233. Observação
    For lI = 1 To 15
        Print #pNumArqDados, " "
    Next

    '234. Local e data.
    lAuxiliar2 = "São Paulo, " & _
                 Format$(Now, "dd") & " de " & Format$(Now, "mmmm") & " de " & Format$(Now, "yyyy")
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '235. Nome do corretor (assinatura)
    lAuxiliar2 = lObjstcCorr.gNom_Corretor
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    mfCriDados_MULTPRO1 = True

End Function

Private Function mfCriDados_MULTPRO2(ByVal pNumArqDados As Integer, ByVal pobjstcPed As stcA46V702B, _
                                     pobjstcPedLoc As stcA46V708B) As Boolean
'Função: cria arquivo de dados - proposta empresarial simples -

'Parâmetro de entrada...pNumArqDados...Número do arquivo de dados.


    Dim lAuxiliar1 As Double  'Auxiliar - utilizado apenas como exemplo.
    Dim lAuxiliar2 As String  'Auxiliar - utilizado apenas como exemplo.
    Dim lAuxiliar_01 As String
    Dim lAuxiliar_02 As String
    Dim lAuxiliar_03 As String
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
    Dim lobjstcPedJur As stcA46V711B
    Dim lSelect As String
    Dim lobjstcClau As stcA46V703B    'Objeto da classe de estrutura stcA46V703B(TAB_PED_CLAU).

    mfCriDados_MULTPRO2 = False

    Dim lobjstcPedLoc As stcA46V708B    'Classe (stc) de itens.

    Dim lbdTab_Tipo_Fator As ADODB.Connection   'Arquivo: P00Tarifa.
    Dim lrsTab_Tipo_Fator As clsYasRecordSet    'Registro :Tab_Tipo_Fator.
    Dim lobjstcFator As stcA46V720B   'Objeto da classe de estrutura stcA46V720B (TAB_PED_FAT_RISCO).


    '001. Título do seguro (de acordo com a base de dados).
    '     Neste caso deve ser colocado ANEXO DA PROPOSTA DE SEGURO YASUDA...
    '     RESIDENCIAL, CONDOMÍNIO ou EMPRESARIAL.
    '****************
    'Print #pNumArqDados, "ANEXO DA PROPOSTA DE SEGURO YASUDA " & UCase(pObjstcPed.gDesc_Ramo)
    '****************
    Print #pNumArqDados, UCase(pobjstcPed.gDesc_Ramo)

    '002. Emissão.
    'Print #pNumArqDados, "Emissão: " & Format$(Now, "dd/mm/yyyy hh:nn:ss")
    Print #pNumArqDados, Format$(Now, "dd/mm/yyyy hh:nn:ss")

    '003. Número do pedido .
    lAuxiliar1 = pobjstcPed.Num_Ped
    Print #pNumArqDados, Format$(lAuxiliar1, "000000")

    '004. Número do item .
    lAuxiliar1 = pobjstcPedLoc.Num_Item
    Print #pNumArqDados, Format$(lAuxiliar1, "0000")

    '005. Corretor (código e nome) .
    If pobjstcPed.ColecaoPedCorr.Count > 0 Then
        Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(1)
        lAuxiliar1 = lObjstcCorr.Cod_Corr
        lAuxiliar2 = lObjstcCorr.gNom_Corretor
    Else
        lAuxiliar1 = 0
        lAuxiliar2 = " "
    End If
    Print #pNumArqDados, Format$(lAuxiliar1, "0000") & " - " & lAuxiliar2

    '006. Tarifa .
    lAuxiliar2 = pobjstcPed.Cod_Produto
    If pobjstcPed.Cod_Produto = 976 And pobjstcPed.Cod_Ramo = e_CodRamo_Empresarial Then
        lAuxiliar2 = lAuxiliar2 & " A"
    End If
    If pobjstcPed.Cod_Sequencia = 1 Then
        lAuxiliar2 = lAuxiliar2 & " (A)"
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '007. Nosso número .
    lAuxiliar2 = pobjstcPed.NOSSO_NUMERO
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    Print #pNumArqDados, pobjstcPed.Cod_Unicidade & " - " & pobjstcPed.Cod_Ver_Unici

    '008 a 012. Declarações/Observações .
    lCountLine = 0
    If pobjstcPed.Cod_Ramo = e_CodRamo_Empresarial Or pobjstcPed.Cod_Ramo = e_CodRamo_Condominio Then
        lCountLine = lCountLine + 1
        lAuxiliar2 = "Valor em risco declarado: R$ " & Format(pobjstcPedLoc.Val_Vr, "standard")
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    End If

    lCountLine = lCountLine + 1
    lAuxiliar2 = "-Existem outros seguros cobrindo os mesmos bens? " & IIf(pobjstcPedLoc.ColecaoPedSeguro.Count > 0, "SIM", "NÃO")
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    lblnAchou = False

    If pobjstcPed.Cod_Ramo <> e_CodRamo_Condominio Then
        lCountLine = lCountLine + 1
        lAuxiliar2 = "-Deseja a inclusão de Cláusula Beneficiária? "

        If frmT46V101A.optBeneficiario(0).Value = True Then
            lAuxiliar2 = lAuxiliar2 & "SIM"
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
    For Each lobjstcPedJur In pobjstcPed.ColecaoPedJuros
        If lobjstcPedJur.Num_Parcela = pobjstcPed.Qtd_Parc Then
            If pobjstcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec = e_SimNao_Sim Then
                pobjstcPed.objstcCtrlEmis.Per_Juros_Neg = lobjstcPedJur.Val_Ant
            Else
                pobjstcPed.objstcCtrlEmis.Per_Juros_Neg = lobjstcPedJur.Val_Pos
            End If
            Exit For
        End If
    Next lobjstcPedJur

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
    Print #pNumArqDados, "FRANQUIAS"
    lCountLine = 1
    'Call mpDescricaoFranquia_Corretor(pobjstcPedLoc, pobjstcPed, lDescFranquias, 52)
    'Não deve mais buscar pois já está gravado
    'Call gM46V999.mpNovaBuscaDescricaoFranquia(pobjstcPed, pobjstcPedLoc, 52, 0, lDescFranquias)

    For lI = 0 To 51
        If lDescFranquias(lI) = "" Then Exit For
        lAuxiliar2 = lDescFranquias(lI)
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        lCountLine = lCountLine + 1
    Next lI

    'Call mpDescricaoFranquia(pobjstcPedLoc, pobjstcPed, lDescFranquias, 52 - lCountLine)
    'For li = 0 To 51 - lCountLine
    '    If lDescFranquias(li) = "" Then Exit For
    '    lAuxiliar2 = lDescFranquias(li)
    '    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    '    lCountLine = lCountLine + 1
    'Next li

    '028 a 037. Importante.
    Print #pNumArqDados, " "
    Print #pNumArqDados, "IMPORTANTE"
    lCountLine = lCountLine + 2

    Call mpTextoImportante(pobjstcPedLoc, pobjstcPed, lDescImportante, "MULTPRO2", 52 - lCountLine)
    For lI = 0 To UBound(lDescImportante) - 1
        If lDescImportante(lI) = "" Then Exit For
        lAuxiliar2 = lDescImportante(lI)
        If gM46V999.gfPreenchido(lAuxiliar2) And lCountLine < 52 Then
            lCountLine = lCountLine + 1
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        End If
    Next lI
    If lCountLine < 52 Then
        Call mpTextoObs(pobjstcPedLoc, lDescImportante, 52 - lCountLine)
        For lI = 0 To UBound(lDescImportante) - 1
            lAuxiliar2 = lDescImportante(lI)
            If lDescImportante(lI) = "" Then Exit For
            If gM46V999.gfPreenchido(lAuxiliar2) Then
                lCountLine = lCountLine + 1
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            End If
        Next lI
    End If

    If (pobjstcPedLoc.Cob_Bens = True) And (pobjstcPedLoc.ColecaoPedLocBem.Count = 0) Then    '1001146 - Tarifa 1000 e 1001 Dezembro 2010
        For Each lobjstcCob In pobjstcPedLoc.ColecaoPedCobert
            If gM46V999.gCod_Ramo = e_CodRamo_Empresarial And lobjstcCob.Cod_Cobert = gM46V999.gfBuscarDePara(94, 1) And pobjstcPedLoc.Dig_Bens = False Then
                lAuxiliar2 = "Para a cobertura " & lobjstcCob.Cod_Cobert & " é obrigatório relacionar os bens e respectivos valores. Opção escolhida pelo corretor: Entrega da relação"
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
                lCountLine = lCountLine + 1
                lAuxiliar2 = "juntamente com a proposta de seguro."
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
                lCountLine = lCountLine + 1
            End If
            If gM46V999.gCod_Ramo = e_CodRamo_Residencial And (lobjstcCob.Cod_Cobert = gM46V999.gfBuscarDePara(27, 1) Or lobjstcCob.Cod_Cobert = gM46V999.gfBuscarDePara(29, 1) Or _
                                                               lobjstcCob.Cod_Cobert = gM46V999.gfBuscarDePara(30, 1) Or lobjstcCob.Cod_Cobert = gM46V999.gfBuscarDePara(31, 1)) And pobjstcPedLoc.Dig_Bens = False Then
                lAuxiliar2 = "Para a cobertura " & lobjstcCob.Cod_Cobert & " é obrigatório relacionar os bens e respectivos valores. Opção escolhida pelo corretor: Entrega da relação"
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
                lCountLine = lCountLine + 1
                lAuxiliar2 = "juntamente com a proposta de seguro."
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
                lCountLine = lCountLine + 1
            End If
        Next lobjstcCob

        GoTo FimTexto
    End If

    ' Questionário

FimTexto:
    If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTab_Tipo_Fator, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Function
    End If
    Dim lFator As Integer
    lFator = 0
    For Each lobjstcPedLoc In pobjstcPed.ColecaoPedLoc
        'Ficha 1000601
        If gM46V999.gCod_Ramo = e_CodRamo_Residencial Then
            If pobjstcPedLoc.gLiberaDescontoEspecial = True Then
                lSelect = " Select distinct DSC_FATOR, COD_FATOR  from Tbm_Tipo_Fator_Protec "
                lSelect = lSelect & " Where cod_Ramo = " & pobjstcPed.Cod_Ramo
                lSelect = lSelect & " and Cod_Modalidade = " & lobjstcPedLoc.Cod_Produto
                lSelect = lSelect & " and COD_TARIF = " & pobjstcPed.Cod_Produto
                lSelect = lSelect & " and Cod_Sequencia <= " & pobjstcPed.Cod_Sequencia
                lSelect = lSelect & " order by DSC_FATOR"
            Else
                lSelect = " Select distinct DSC_FATOR, COD_FATOR  from Tbm_Tipo_Fator_Protec "
                lSelect = lSelect & " Where cod_Ramo = " & pobjstcPed.Cod_Ramo
                lSelect = lSelect & " and Cod_Modalidade = " & lobjstcPedLoc.Cod_Produto
                lSelect = lSelect & " and COD_TARIF = " & pobjstcPed.Cod_Produto
                lSelect = lSelect & " and Cod_Sequencia <= " & pobjstcPed.Cod_Sequencia
                lSelect = lSelect & " and Cod_Fator <> 7 "    'Não carregar o item Desconto Especial
                lSelect = lSelect & " order by DSC_FATOR"
            End If
        Else
            lSelect = " Select distinct DSC_FATOR, COD_FATOR  from Tbm_Tipo_Fator_Protec "
            lSelect = lSelect & " Where cod_Ramo = " & pobjstcPed.Cod_Ramo
            lSelect = lSelect & " and Cod_Modalidade = " & lobjstcPedLoc.Cod_Produto
            lSelect = lSelect & " and COD_TARIF = " & pobjstcPed.Cod_Produto
            lSelect = lSelect & " and Cod_Sequencia <= " & pobjstcPed.Cod_Sequencia
            lSelect = lSelect & " order by DSC_FATOR"
        End If

        If gM46V999.gfObtRegistro(lbdTab_Tipo_Fator, lSelect, lrsTab_Tipo_Fator, lMensagem) = False Then
            Call gM46V999.gpFecha2(lbdTab_Tipo_Fator)
            Call gM46V999.gpGraLog(2, lMensagem)
            Screen.MousePointer = vbDefault
            Exit Function
        End If
        If Not lrsTab_Tipo_Fator.EOF Then
            Print #pNumArqDados, " "
            Print #pNumArqDados, "QUESTIONÁRIO"
            lCountLine = lCountLine + 2
        End If

        Do While Not lrsTab_Tipo_Fator.EOF
            lCountLine = lCountLine + 1
            lFator = lFator + 1
            lblnAchou = False

            For Each lobjstcFator In lobjstcPedLoc.ColecaoPedFator
                If lobjstcFator.Cod_Fator = lrsTab_Tipo_Fator("Cod_Fator") Then
                    lblnAchou = True
                    lAuxiliar2 = Format(lFator, "00") & " (X) " & Trim(lrsTab_Tipo_Fator("Dsc_Fator"))
                    If Trim(lobjstcFator.Dsc_Compl_Fator) <> "" Then
                        lAuxiliar2 = lAuxiliar2 & " - " & Trim(lobjstcFator.Dsc_Compl_Fator)
                    End If
                    Print #pNumArqDados, lAuxiliar2
                    Exit For
                End If
            Next lobjstcFator

            If lblnAchou = False Then
                lAuxiliar2 = Format(lFator, "00") & " ( ) " & Trim(lrsTab_Tipo_Fator("Dsc_Fator"))
                Print #pNumArqDados, lAuxiliar2
            End If
            lrsTab_Tipo_Fator.MoveNext
        Loop
    Next

    Call gM46V999.gpFecha2(lbdTab_Tipo_Fator)

    Do While lCountLine < 52
        lCountLine = lCountLine + 1
        Print #pNumArqDados, " "
    Loop


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

    'marcia
    If pobjstcPed.Cod_Unid_Prod < 1000 Then
        lAuxiliar2 = "MATRIZ"
    End If

    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '053. Produtor .
    If lAuxiliar2 = "MATRIZ" Then
        lAuxiliar2 = ""
    Else
        lAuxiliar2 = pobjstcPed.Cod_Prod
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    'marcia

    '054. Corretor .
    lAuxiliar2 = " "
    'If pObjstcPed.ColecaoPedCorr.Count > 0 Then
    'Set lObjstcCorr = pObjstcPed.ColecaoPedCorr.Item(1)
    'lAuxiliar2 = Format(lObjstcCorr.COD_CORR, "0000000")
    'If pObjstcPed.ColecaoPedCorr.Count > 1 Then
    'Set lObjstcCorr = pObjstcPed.ColecaoPedCorr.Item(2)
    'lAuxiliar2 = lAuxiliar2 & " " & Format(lObjstcCorr.COD_CORR, "0000000")
    'If pObjstcPed.ColecaoPedCorr.Count > 2 Then
    'Set lObjstcCorr = pObjstcPed.ColecaoPedCorr.Item(3)
    'lAuxiliar2 = lAuxiliar2 & " " & Format(lObjstcCorr.COD_CORR, "0000000")
    'Else
    'lAuxiliar2 = lAuxiliar2 & " 0000000 "
    'End If
    'Else
    'lAuxiliar2 = lAuxiliar2 & " 0000000 0000000 "
    'End If
    'End If

    'Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '###########################################################################

    lAuxiliar_01 = ""
    lAuxiliar_02 = ""
    lAuxiliar_03 = ""

    If pobjstcPed.ColecaoPedCorr.Count > 0 Then
        Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(1)
        lAuxiliar_01 = Format(lObjstcCorr.Cod_Corr, "0000000")
        lAuxiliar_01 = lAuxiliar_01 & "    C " & Format(lObjstcCorr.Per_Corr * 100, "0000")

        If pobjstcPed.ColecaoPedCorr.Count > 1 Then
            Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(2)
            lAuxiliar_02 = Format(lObjstcCorr.Cod_Corr, "0000000")
            lAuxiliar_02 = lAuxiliar_02 & "    C " & Format(lObjstcCorr.Per_Corr * 100, "0000")

            If pobjstcPed.ColecaoPedCorr.Count > 2 Then
                Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(3)
                lAuxiliar_03 = Format(lObjstcCorr.Cod_Corr, "0000000")
                lAuxiliar_03 = lAuxiliar_03 & "    C " & Format(lObjstcCorr.Per_Corr * 100, "0000")

            Else
                lAuxiliar_03 = " "

            End If

        Else
            lAuxiliar_02 = " "
            lAuxiliar_03 = " "

        End If

    Else
        lAuxiliar_01 = " "
        lAuxiliar_02 = " "
        lAuxiliar_03 = " "

    End If

    Print #pNumArqDados, lAuxiliar_01
    Print #pNumArqDados, lAuxiliar_02
    Print #pNumArqDados, lAuxiliar_03

    '###########################################################################
    '055. Informação em código .

    'lAuxiliar2 = " "
    'If pObjstcPed.ColecaoPedCorr.Count > 0 Then
    '    Set lObjstcCorr = pObjstcPed.ColecaoPedCorr.Item(1)
    '    lAuxiliar2 = "C " & lObjstcCorr.Per_Corr * 100
    '    If pObjstcPed.ColecaoPedCorr.Count > 1 Then
    '        Set lObjstcCorr = pObjstcPed.ColecaoPedCorr.Item(2)
    '        lAuxiliar2 = lAuxiliar2 & " " & lObjstcCorr.Per_Corr * 100
    '        If pObjstcPed.ColecaoPedCorr.Count > 2 Then
    '            Set lObjstcCorr = pObjstcPed.ColecaoPedCorr.Item(3)
    '            lAuxiliar2 = lAuxiliar2 & " " & lObjstcCorr.Per_Corr * 100
    '        End If
    '    End If
    'End If

    'Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

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

    mfCriDados_MULTPRO2 = True

End Function
Private Function mfCriDados_MULTPRO3(ByVal pNumArqDados As Integer, ByVal pobjstcPed As stcA46V702B, _
                                     pobjstcPedLoc As stcA46V708B, pTotalPags As Integer) As Boolean
'Função: cria arquivo de dados - proposta coletiva - capa.

'Parâmetro de entrada...pNumArqDados...Número do arquivo de dados.

    Dim lAuxiliar1 As Double  'Auxiliar - utilizado apenas como exemplo.
    Dim lAuxiliar2 As String  'Auxiliar - utilizado apenas como exemplo.
    Dim lAuxiliar3 As String  'Auxiliar - utilizado apenas como exemplo.
    Dim lAuxiliar_01 As String
    Dim lAuxiliar_02 As String
    Dim lAuxiliar_03 As String

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
    Dim lobjstcPedJur As stcA46V711B
    mfCriDados_MULTPRO3 = False

    '001. Título do documento.
    Print #pNumArqDados, UCase(pobjstcPed.gDesc_Ramo)

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
    'Print #pNumArqDados, "Emissão: " & Format$(Now, "dd/mm/yyyy hh:nn:ss")
    Print #pNumArqDados, Format$(Now, "dd/mm/yyyy hh:nn:ss")

    '004. Número da página.
    'Print #pNumArqDados, "Página 1/" & pTotalPags
    Print #pNumArqDados, "1/" & pTotalPags

    '005. Número do pedido .
    lAuxiliar1 = pobjstcPed.Num_Ped
    Print #pNumArqDados, Format$(lAuxiliar1, "000000")

    '006. Corretor (código e nome) .
    If pobjstcPed.ColecaoPedCorr.Count > 0 Then
        Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(1)
        lAuxiliar1 = lObjstcCorr.Cod_Corr
        lAuxiliar2 = lObjstcCorr.gNom_Corretor
    Else
        lAuxiliar1 = 0
        lAuxiliar2 = " "
    End If
    Print #pNumArqDados, Format$(lAuxiliar1, "0000") & " - " & lAuxiliar2

    '007. Tarifa .

    lAuxiliar_02 = pobjstcPed.Cod_Produto
    If pobjstcPed.Cod_Produto = 976 And pobjstcPed.Cod_Ramo = e_CodRamo_Empresarial Then
        lAuxiliar2 = lAuxiliar2 & " A"
    End If
    If pobjstcPed.Cod_Sequencia = 1 Then
        lAuxiliar2 = lAuxiliar2 & " (A)"
    End If
    Print #pNumArqDados, lAuxiliar_02
    '008. Nosso número .
    Print #pNumArqDados, pobjstcPed.NOSSO_NUMERO


    If pobjstcPed.objstcCtrlEmis.Tip_Cotac = 2 And pobjstcPed.objstcCtrlEmis.Num_Cotac <> 0 Then
        Print #pNumArqDados, "Siscota: " & pobjstcPed.objstcCtrlEmis.Num_Cotac
    Else
        Print #pNumArqDados, " "
    End If


    Print #pNumArqDados, pobjstcPed.Cod_Unicidade & " - " & pobjstcPed.Cod_Ver_Unici

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
        'lAuxiliar2 = "CNPJ.........:"
        lAuxiliar2 = "              "
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        lAuxiliar2 = pobjstcPed.gCNPJ_CPF_Formatado & "   " & _
                     "Ramo de atividade: " & pobjstcPed.gDesc_RamoAtiv

        If pobjstcPed.Val_RendaFatu > 0 Then
            lAuxiliar2 = lAuxiliar2 & " Faturamento: " & Format(pobjstcPed.Val_RendaFatu, "standard")
        Else
            lAuxiliar2 = lAuxiliar2 & " Faturamento: não informada "
        End If
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    Case e_TipPessoa_Fisica
        Print #pNumArqDados, "Física"
        '012 e 013. Proponente - CNPJ ou CPF .
        'lAuxiliar2 = "CPF...........:"
        lAuxiliar2 = "               "
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        lAuxiliar2 = pobjstcPed.gCNPJ_CPF_Formatado & "   " & _
                     pobjstcPed.gCod_Doc_Segur & ": " & Trim(pobjstcPed.Num_Doc) & "   " & _
                     "Órgão emissor: " & Trim(pobjstcPed.Org_Exp) & "   " & _
                     "Emissão: " & gM46V999.gfForData(2, pobjstcPed.Dat_Emis_Doc)
        If pobjstcPed.Val_RendaFatu > 0 Then
            lAuxiliar2 = lAuxiliar2 & " Renda: " & Format(pobjstcPed.Val_RendaFatu, "standard")
        Else
            lAuxiliar2 = lAuxiliar2 & " Renda: não informada "
        End If
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    End Select

    '014. Proponente - endereço .
    lAuxiliar2 = pobjstcPed.objstcPedEndereco.Cod_Tip_Logr & " " & _
                 pobjstcPed.objstcPedEndereco.Nom_Logr & "," & _
                 pobjstcPed.objstcPedEndereco.Dsc_Compl_Logr
    'Ficha 900417
    Do While Len(lAuxiliar2) < 70
        lAuxiliar2 = lAuxiliar2 & " "
    Loop

    If gM46V999.gCodProduto >= 968 Then
        lAuxiliar2 = lAuxiliar2 & "Profissão: " & pobjstcPed.Nom_Empresa
    Else
        lAuxiliar2 = lAuxiliar2 & "Empresa: " & pobjstcPed.Nom_Empresa
    End If
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

    '023a. Resumo: PER IOF.
    lAuxiliar1 = gM46V111.gProJurIOF
    lAuxiliar2 = " IOF (" & Format$(lAuxiliar1, "standard") & "%):"
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
    For Each lobjstcPedJur In pobjstcPed.ColecaoPedJuros
        If lobjstcPedJur.Num_Parcela = pobjstcPed.Qtd_Parc Then
            If pobjstcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec = e_SimNao_Sim Then
                pobjstcPed.objstcCtrlEmis.Per_Juros_Neg = lobjstcPedJur.Val_Ant
            Else
                pobjstcPed.objstcCtrlEmis.Per_Juros_Neg = lobjstcPedJur.Val_Pos
            End If
            Exit For
        End If
    Next lobjstcPedJur
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

    ''032 a 035. Observação 1.
    ''a) Linha 1.
    'lAuxiliar2 = "A aceitação do seguro estará sujeita à análise do risco. O registro deste plano na " & _
     '             "SUSEP não implica,  por  parte"
    'Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    ''b) Linha 2.
    'lAuxiliar2 = "da Autarquia, incentivo ou recomendação à sua comercialização. O segurado poderá " & _
     '             "consultar a  situação  cadastral"
    'Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    ''c) Linha 3.
    'lAuxiliar2 = "do seu corretor de seguros no site www.susep.gov.br, por meio do número de seu " & _
     '             "registro na SUSEP, nome  completo, "
    'Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    ''d) Linha 4.
    'lAuxiliar2 = "CNPJ ou CPF."
    'Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    '
    ''036 a 047. Observação 2.
    ''a) Linha 1.
    'lAuxiliar2 = "Propõe(m) à YASUDA SEGUROS S.A. a realização do seguro, declarando conhecer " & _
     '             "previamente e  aceitar  as  Condições"
    'Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    ''b) Linha 2.
    'lAuxiliar2 = "Gerais do(s) anexo(s) que servem de base para esta proposta. Também presta(m)  as  " & _
     '             "informações  solicitadas  pela"
    'Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    ''c) Linha 3.
    'lAuxiliar2 = "YASUDA SEGUROS S.A., constantes do(s) anexo(s) a esta, para  fins  de  estudos  " & _
     '             "da  concretização  do  respectivo"
    'Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    ''d) Linha 4.
    'lAuxiliar2 = "contrato.  Assume(m)  o(s)  proponente(s),  nesta  oportunidade,  integral  " & _
     '             "responsabilidade  pela  exatidão  das"
    'Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    ''e) Linha 5.
    'lAuxiliar2 = "informações ora prestadas, autorizando a YASUDA SEGUROS S.A., caso aceite esta " & _
     '             "proposta, a  emitir  a  respectiva"
    'Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    ''f) Linha 6.
    'lAuxiliar2 = "apólice, cujo prêmio e eventuais despesas o(s) proponente(s) compromete(m)-se  a  " & _
     '             "pagar,  tão  logo  lhe(s)  seja"
    'Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    ''g) Linha 7.
    'lAuxiliar2 = "exigido. Em caso de recusa da proposta de seguro, os valores eventualmente " & _
     '             "adiantados para  futuro  pagamento  de"
    'Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    ''h) Linha 8.
    'lAuxiliar2 = "prêmio pelo(s) proponente(s) serão devolvidos integralmente, no prazo de 10 dias " & _
     '             "corridos da data de sua  recusa."
    'Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    ''i) Linha 9.
    'lAuxiliar2 = "Ultrapassado este prazo, os valores sujeitam-se a atualização monetária pelo " & _
     '             "IPC/FIPE, contados a partir da  data"
    'Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    ''j) Linha 10.
    'lAuxiliar2 = "em que se tornarem exigíveis e juros moratórios de 0,5% (cinco décimos por cento) " & _
     '             "ao mês, contados  a  partir  do"
    'Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    ''k) Linha 11.
    'lAuxiliar2 = "primeiro dia posterior ao término do prazo fixado acima. Ambos calculados " & _
     '             "pro-rata-die  até  a  data  da  efetiva"
    'Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    ''l) Linha 12.
    'lAuxiliar2 = "restituição."
    'Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

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
    lAuxiliar_01 = ""
    lAuxiliar_02 = ""
    lAuxiliar_03 = ""

    If pobjstcPed.ColecaoPedCorr.Count > 0 Then
        Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(1)
        lAuxiliar_01 = Format(lObjstcCorr.Cod_Corr, "0000000")
        lAuxiliar_01 = lAuxiliar_01 & "    C " & Format(lObjstcCorr.Per_Corr * 100, "0000")

        If pobjstcPed.ColecaoPedCorr.Count > 1 Then
            Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(2)
            lAuxiliar_02 = Format(lObjstcCorr.Cod_Corr, "0000000")
            lAuxiliar_02 = lAuxiliar_02 & "    C " & Format(lObjstcCorr.Per_Corr * 100, "0000")

            If pobjstcPed.ColecaoPedCorr.Count > 2 Then
                Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(3)
                lAuxiliar_03 = Format(lObjstcCorr.Cod_Corr, "0000000")
                lAuxiliar_03 = lAuxiliar_03 & "    C " & Format(lObjstcCorr.Per_Corr * 100, "0000")

            Else
                lAuxiliar_03 = " "

            End If
        Else
            lAuxiliar_02 = " "
            lAuxiliar_03 = " "

        End If

    Else
        lAuxiliar_01 = " "
        lAuxiliar_02 = " "
        lAuxiliar_03 = " "

    End If

    Print #pNumArqDados, lAuxiliar_01
    Print #pNumArqDados, lAuxiliar_02
    Print #pNumArqDados, lAuxiliar_03

    '052. Informação em código .
    Print #pNumArqDados, " "


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

    mfCriDados_MULTPRO3 = True
End Function


Private Function mfCriDados_MULTPRO4(ByVal pNumArqDados As Integer, ByVal pobjstcPed As stcA46V702B, _
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
    Dim lPrazo As Integer
    Dim lCodCobert As Integer
    Dim lNumMeses As Integer
    Dim lDescCom As Double
    Dim lDescExp As Double
    Dim lbdTbm_Regiao As ADODB.Connection
    Dim lrsTbm_Regiao As clsYasRecordSet
    Dim lMensagem As String
    Dim lSelect As String
    Dim lDescFranquias() As String
    Dim lDescImportante() As String
    Dim lobjstcClau As stcA46V703B    'Objeto da classe de estrutura stcA46V703B(TAB_PED_CLAU).
    Dim lblnAchou As Boolean

    mfCriDados_MULTPRO4 = False

    lQuant = 1
    For Each lobjstcPedLoc In pobjstcPed.ColecaoPedLoc
        '001 e 002. Título do documento.
        '*************************
        'Print #pNumArqDados, "ANEXO DA PROPOSTA DE SEGURO YASUDA"
        '*************************
        Print #pNumArqDados, UCase(pobjstcPed.gDesc_Ramo)

        '003. Emissão.
        '*************************
        'Print #pNumArqDados, "Emissão: " & Format$(Now, "dd/mm/yyyy hh:nn:ss")
        lAuxiliar2 = Format$(Now, "dd/mm/yyyy hh:nn:ss")
        If pobjstcPed.Num_PI_Emissao <> 0 Then
            lAuxiliar2 = lAuxiliar2 & Space(20) & "PI: " & pobjstcPed.Num_PI_Emissao & "-" & pobjstcPed.Dig_PI_Emissao
        End If
        If pobjstcPed.Ind_Reentrada = 1 Then
            lAuxiliar2 = lAuxiliar2 & "**REENTRADA**"
        End If
        Print #pNumArqDados, lAuxiliar2
        '*************************
        lQuant = lQuant + 1
        '004. Número da página.
        '*************************
        'Print #pNumArqDados, "Página " & lQuant & "/" & pTotalPag
        Print #pNumArqDados, lQuant & "/" & pTotalPag
        '*************************
        '005. Número do pedido .
        Print #pNumArqDados, pobjstcPed.Num_Ped

        '006. Número do item .
        Print #pNumArqDados, lobjstcPedLoc.Num_Item

        '007. Corretor (código e nome) .
        If pobjstcPed.ColecaoPedCorr.Count > 0 Then
            Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(1)
            lAuxiliar1 = lObjstcCorr.Cod_Corr
            lAuxiliar2 = lObjstcCorr.gNom_Corretor
        Else
            lAuxiliar1 = 0
            lAuxiliar2 = " "
        End If
        Print #pNumArqDados, Format$(lAuxiliar1, "0000") & " - " & lAuxiliar2


        '008. Tarifa .
        lAuxiliar2 = pobjstcPed.Cod_Produto
        If pobjstcPed.Cod_Produto = 976 And pobjstcPed.Cod_Ramo = e_CodRamo_Empresarial Then
            lAuxiliar2 = lAuxiliar2 & " A"
        End If
        If pobjstcPed.Cod_Sequencia = 1 Then
            lAuxiliar2 = lAuxiliar2 & " (A)"
        End If
        Print #pNumArqDados, lAuxiliar2

        '009. Nosso número .
        Print #pNumArqDados, pobjstcPed.NOSSO_NUMERO
        Print #pNumArqDados, pobjstcPed.Cod_Unicidade & " - " & pobjstcPed.Cod_Ver_Unici

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
        If gM46V999.gCod_Ramo = e_CodRamo_Empresarial Then
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
        'lAuxiliar2 = lObjstcPedLoc.Cod_Regiao & " - " & Left(lObjstcPedLoc.gDesc_Regiao, 38)
        'Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        If pobjstcPed.Cod_Produto <= gM46V999.TarifaAntiga Then
            lAuxiliar2 = lobjstcPedLoc.Cod_Regiao & " - " & Left(lobjstcPedLoc.gDesc_Regiao, 38)
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        Else
            'Nova Região
            If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTbm_Regiao, lMensagem) = False Then
                Call gM46V999.gpGraLog(2, lMensagem)
                Exit Function
            End If

            lSelect = "Select Cod_Faixa_Reg, DSC_REG from Tbm_Faixa_Regiao "
            lSelect = lSelect & " Where cod_Ramo = " & pobjstcPed.Cod_Ramo
            lSelect = lSelect & "    and Cod_Modalidade = " & lobjstcPedLoc.Cod_Produto
            lSelect = lSelect & "    and Cod_Tarif = " & pobjstcPed.Cod_Produto
            lSelect = lSelect & "    and COD_Sequencia <= " & pobjstcPed.Cod_Sequencia
            lSelect = lSelect & "    and Cep_Ini <= " & lobjstcPedLoc.Num_Cep_Risco
            lSelect = lSelect & "    and Cep_Fim >= " & lobjstcPedLoc.Num_Cep_Risco
            lSelect = lSelect & "    order by Cod_Sequencia desc "

            If gM46V999.gfObtRegistro(lbdTbm_Regiao, lSelect, lrsTbm_Regiao, lMensagem) = False Then
                Call gM46V999.gpGraLog(0, lMensagem)
                Call gM46V999.gpFecha2(lbdTbm_Regiao)
                Exit Function
            End If

            If Not lrsTbm_Regiao.EOF Then
                lAuxiliar2 = gM46V999.Check(lrsTbm_Regiao("Cod_Faixa_Reg")) & " - " & Left(gM46V999.Check(lrsTbm_Regiao("DSC_REG")), 38)
            Else
                lAuxiliar2 = "Não cadastrada."
            End If
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        End If


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
        Case e_TipEmissao_RenYasuda, e_TipEmissao_Endosso
            lAuxiliar2 = lAuxiliar2 & "  Apólice: " & lobjstcPedLoc.Num_Apol_Anterior & "  Item: " & lobjstcPedLoc.Num_Item_Anterior & "  Sinistro: " & IIf(lobjstcPedLoc.Ind_Sinistro = e_SimNao_Sim, "Sim", "Não")
        End Select
        Print #pNumArqDados, lAuxiliar2

        '025 e 026. Coberturas contratadas - desconto - título.
        'm       If pObjstcPed.Cod_Ramo = e_CodRamo_Empresarial Then
        'm           '035 e 036. Coberturas contratadas - título.
        'm          lAuxiliar2 = "Desconto de"
        'm           Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'm           lAuxiliar2 = "experiência"
        'm           Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'm       Else
        Print #pNumArqDados, " "
        Print #pNumArqDados, " "
        'm      End If


        '027 a 158. Coberturas contratadas - linhas de 01 a 22.
        lCountLine = 0
        lValTotCober = 0
        For Each lobjstcCob In lobjstcPedLoc.ColecaoPedCobert
            lCountLine = lCountLine + 1
            'Código
            Print #pNumArqDados, Format$(lobjstcCob.Cod_Cobert, "0000")
            lCodCobert = lobjstcCob.Cod_Cobert
            'Descrição
            Print #pNumArqDados, Left(lobjstcCob.gNomeCobertura, 45)
            'm
            lAuxiliar1 = IIf(IsNumeric(lobjstcCob.Num_Meses_PI), lobjstcCob.Num_Meses_PI, 0)
            If lAuxiliar1 = 0 Then
                Print #pNumArqDados, " "
                Print #pNumArqDados, " "
            Else
                'pi mes
                Print #pNumArqDados, Format$(lobjstcCob.Num_Meses_PI, "00")
                lNumMeses = lobjstcCob.Num_Meses_PI
                Call mpPrazoTabCoefPI(pobjstcPed, lCodCobert, lNumMeses, lPrazo)
                'pi dias
                If lPrazo <> 0 Then
                    Print #pNumArqDados, lPrazo
                Else
                    Print #pNumArqDados, " "
                End If
            End If
            '%Sobre Prejuízos
            lAuxiliar1 = lobjstcCob.Per_Franq
            lAuxiliar2 = Format$(lAuxiliar1, "standard")
            lTamanho = Len(lAuxiliar2)
            If lTamanho < 10 Then
                lAuxiliar2 = Space(10 - lTamanho) & lAuxiliar2
            End If
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

            'Limitado ao Mínimo
            lAuxiliar1 = lobjstcCob.Val_Franq_Min
            lAuxiliar2 = Format$(lAuxiliar1, "standard")
            lTamanho = Len(lAuxiliar2)
            If lTamanho < 14 Then
                lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
            End If
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            'm
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
            lDescCom = lobjstcCob.Per_Desc_Coml
            lAuxiliar2 = Format$(lAuxiliar1, "standard")
            lTamanho = Len(lAuxiliar2)
            If lTamanho < 10 Then
                lAuxiliar2 = Space(10 - lTamanho) & lAuxiliar2
            End If
            'm            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

            If pobjstcPed.Cod_Ramo = e_CodRamo_Empresarial Then
                'e) Desconto de experiência .
                lAuxiliar1 = lobjstcCob.Per_Desc_Exp
                lAuxiliar2 = Format$(lAuxiliar1, "standard")
                lTamanho = Len(lAuxiliar2)
                If lTamanho < 10 Then
                    lAuxiliar2 = Space(10 - lTamanho) & lAuxiliar2
                End If
                'm                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            Else
                lDescExp = 0
                'm                Print #pNumArqDados, " "
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
        If lValTotCober < gM46V999.gProPremioMinimo Then
            lValTotCober = gM46V999.gProPremioMinimo
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

        '*  lAuxiliar2 = lAuxiliar2 & "/A:" & Format(lObjstcPedLoc.Per_Desc_Apar, "00")

        'Ficha 900740 e 901673 e 901772
        If (gM46V111.gCod_plano >= 2021 And gM46V111.gCod_plano <= 2031) Or (gM46V111.gCod_plano >= 2061 And gM46V111.gCod_plano <= 2100) Or (gM46V111.gCod_plano >= 2111 And gM46V111.gCod_plano <= 2130) Then
            lAuxiliar2 = lAuxiliar2 & "/P:00"
        Else
            lAuxiliar2 = lAuxiliar2 & "/P:" & Format(lobjstcPedLoc.Per_Desc_Progres, "00")
        End If

        'lAuxiliar2 = lAuxiliar2 & "/CO:" & Format(pObjstcPed.objstcCtrlEmis.Per_Corr + pObjstcPed.objstcCtrlEmis.Per_Corr_2 + pObjstcPed.objstcCtrlEmis.Per_Corr_3, "00.00")
        lAuxiliar1 = 0
        If pobjstcPed.ColecaoPedCorr.Count > 0 Then
            Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(1)
            lAuxiliar1 = lObjstcCorr.Per_Corr
            If pobjstcPed.ColecaoPedCorr.Count > 1 Then
                Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(2)
                lAuxiliar1 = lAuxiliar1 + lObjstcCorr.Per_Corr
                If pobjstcPed.ColecaoPedCorr.Count > 2 Then
                    Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(3)
                    lAuxiliar1 = lAuxiliar1 + lObjstcCorr.Per_Corr
                End If
            End If
        End If
        lAuxiliar2 = lAuxiliar2 & "/CO:" & Format(lAuxiliar1 * 100, "0000")
        '*   lAuxiliar2 = lAuxiliar2 & "/C:" & Format(Round(lObjstcPedLoc.Val_Coef_IS_Vr, 4) * 10000, "00000")

        'Ficha 900740 e 901673 e 901772
        If (gM46V111.gCod_plano >= 2021 And gM46V111.gCod_plano <= 2031) Or (gM46V111.gCod_plano >= 2061 And gM46V111.gCod_plano <= 2100) Or (gM46V111.gCod_plano >= 2111 And gM46V111.gCod_plano <= 2130) Then
            lAuxiliar2 = lAuxiliar2 & "/DC:00"
        Else
            lAuxiliar2 = lAuxiliar2 & "/DC:" & Format(lDescCom * 100, "0000")
        End If

        '*   lAuxiliar2 = lAuxiliar2 & "/DE:" & Format(lDescExp, "00")
        '   lAuxiliar2 = lAuxiliar2 & "/PRGRS6"

        'Ficha 900550
        '/V1.0 -> versão após ajustar o bônus das coberturas

        lAuxiliar2 = lAuxiliar2 & "/V1.0"

        '###INSPECAO
        If gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Or gM46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda Then
            If lobjstcPedLoc.Insp_Obr = True Then
                If lobjstcPedLoc.Insp_Agend = True Then
                    lAuxiliar2 = lAuxiliar2 & "/INSPS"
                Else
                    lAuxiliar2 = lAuxiliar2 & "/INSPN"
                End If
            End If
        End If

        lAuxiliar2 = lAuxiliar2 & "/" & Format(gM46V111.gProJurIOF * 100, "000")

        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        'm        '162 a 171. Segurança.
        'm        '           O texto pode ser customizado de acordo com a necessidade.
        'm        '           Temos 10 linhas disponíveis e cada linha cabe 70 caracteres (contando com o "(X)").
        'm        '           O "X" deve ser colocado de acordo com a base de dados.
        'm        If pObjstcPed.Cod_Ramo = e_CodRamo_Condominio Then
        'm            'a) Linha 01.
        'm            Print #pNumArqDados, "SEGURANÇA"
        'm            If lobjstcPedLoc.Cod_Iden_Segur_01_V = "S" Then
        'm                lAuxiliar2 = "(S) "
        'm            Else
        'm                lAuxiliar2 = "(N) "
        'm            End If
        'm            lAuxiliar2 = lAuxiliar2 & "Vigilância exclusiva e permanente do imóvel entendendo-se ainda"
        'm            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'm            'b) Linha 02.
        'm            lAuxiliar2 = "    como tal, porteiros fixos."
        'm            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'm            'c) Linha 03.
        'm            If lobjstcPedLoc.Cod_Iden_Segur_02_A = "S" Then
        'm                lAuxiliar2 = "(S) "
        'm            Else
        'm                lAuxiliar2 = "(N) "
        'm            End If
        'm            lAuxiliar2 = lAuxiliar2 & "Alarmes."
        'm            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'm            'd) Linha 04.
        'm            If lobjstcPedLoc.Cod_Iden_Segur_03_P = "S" Then
        'm                lAuxiliar2 = "(S) "
        'm            Else
        'm                lAuxiliar2 = "(N) "
        'm            End If
        'm            lAuxiliar2 = lAuxiliar2 & "Porteiros eletrônicos."
        'm            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'm            'e) Linha 05.
        'm            If lobjstcPedLoc.Cod_Iden_Segur_04_I = "S" Then
        'm                lAuxiliar2 = "(S) "
        'm            Else
        'm                lAuxiliar2 = "(N) "
        'm            End If
        'm            lAuxiliar2 = lAuxiliar2 & "Inexistência de terreno baldio, imóvel desocupado ou em construção"
        'm            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'm            'f) Linha 06.
        'm            lAuxiliar2 = "    na vizinhança do risco."
        'm            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'm            'g) Linha 07.
        'm            If lobjstcPedLoc.Cod_Iden_Segur_05_G = "S" Then
        'm                lAuxiliar2 = "(S) "
        'm            Else
        'm                lAuxiliar2 = "(N) "
        'm            End If
        'm            lAuxiliar2 = lAuxiliar2 & "Grades metálicas de proteção em todas as janelas e portas de vidro"
        'm            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'm            'h) Linha 08.
        'm            lAuxiliar2 = "    que dão acesso ao interior do imóvel desde que também possuam"
        'm            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'm            'i) Linha 09.
        'm            lAuxiliar2 = "    alarme."
        'm            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'm            'j) Linha 10.
        'm            lAuxiliar2 = " "
        'm            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'm        Else
        'm            Print #pNumArqDados, " "
        'm            For lI = 1 To 10
        'm                Print #pNumArqDados, String(71, "X")
        'm            Next lI
        'm        End If

        '194 a 198. Proteção.
        '           Para residencial, sairá em branco.
        'm        If pObjstcPed.Cod_Ramo = e_CodRamo_Condominio Or pObjstcPed.Cod_Ramo = e_CodRamo_Empresarial Then
        'm            '           O texto pode ser customizado de acordo com a necessidade.
        'm            '           Temos 4 linhas disponíveis e cada linha cabe 39 caracteres (contando com o "(X)").
        'm            '           O "X" deve ser colocado de acordo com a base de dados.
        'm            'a) Título.
        'm            lAuxiliar2 = "PROTEÇÃO"
        'm            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'm            'b) Linha 01.
        'm            If lobjstcPedLoc.Cod_Iden_Protec_01_E = "S" Then
        'm                lAuxiliar2 = "(S) "
        'm            Else
        'm                lAuxiliar2 = "(N) "
        'm            End If
        'm            lAuxiliar2 = lAuxiliar2 & "Extintor"
        'm            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'm            'c) Linha 02.
        'm            If lobjstcPedLoc.Cod_Iden_Protec_02_H = "S" Then
        'm                lAuxiliar2 = "(S) "
        'm            Else
        'm                lAuxiliar2 = "(N) "
        'm            End If
        'm            lAuxiliar2 = lAuxiliar2 & "Hidrantes"
        'm            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'm            'd) Linha 03.
        'm            If lobjstcPedLoc.Cod_Iden_Protec_03_S = "S" Then
        'm                lAuxiliar2 = "(S) "
        'm            Else
        'm                lAuxiliar2 = "(N) "
        'm            End If
        'm            lAuxiliar2 = lAuxiliar2 & "Sprinkler (chuveiros automáticos)"
        'm            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'm            'e) Linha 04.
        'm            If lobjstcPedLoc.Cod_Iden_Protec_04_D = "S" Then
        'm                lAuxiliar2 = "(S) "
        'm            Else
        'm                lAuxiliar2 = "(N) "
        'm            End If
        'm            lAuxiliar2 = lAuxiliar2 & "Detecção e alarme"
        'm            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        'm        Else
        'm            Print #pNumArqDados, " "
        'm            For lI = 1 To 4
        'm                Print #pNumArqDados, String(40, "X")
        'm            Next lI
        'm        End If

        '177 a 181. Declarações/Observações .
        lCountLine = 0
        If pobjstcPed.Cod_Ramo = e_CodRamo_Empresarial Or pobjstcPed.Cod_Ramo = e_CodRamo_Condominio Then
            lCountLine = lCountLine + 1
            lAuxiliar2 = "Valor em risco declarado: R$ " & Format(lobjstcPedLoc.Val_Vr, "standard")
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        End If

        lCountLine = lCountLine + 1
        lAuxiliar2 = "Existem outros seguros cobrindo os mesmos bens? " & IIf(lobjstcPedLoc.ColecaoPedSeguro.Count > 0, "SIM", "NÃO")
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        lblnAchou = False
        If pobjstcPed.Cod_Ramo <> e_CodRamo_Condominio Then
            lCountLine = lCountLine + 1
            lAuxiliar2 = "Deseja a inclusão de Cláusula Beneficiária? "

            If frmT46V101A.optBeneficiario(0).Value = True Then
                lAuxiliar2 = lAuxiliar2 & "SIM"
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
        Print #pNumArqDados, "FRANQUIAS"
        lCountLine = 1
        'Call mpDescricaoFranquia_Corretor(lObjstcPedLoc, pobjstcPed, lDescFranquias, 28)

        'Não deve mais buscar pois já está gravado
        'Call gM46V999.mpNovaBuscaDescricaoFranquia(pobjstcPed, lObjstcPedLoc, 28, 0, lDescFranquias)

        For lI = 0 To 50
            If lDescFranquias(lI) = "" Then Exit For
            lAuxiliar2 = lDescFranquias(lI)
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            lCountLine = lCountLine + 1
        Next lI

        'Call mpDescricaoFranquia(lobjstcPedLoc, pobjstcPed, lDescFranquias, 28 - lCountLine)
        'For li = 0 To 27 - lCountLine
        '    If lDescFranquias(li) = "" Then Exit For
        '    lAuxiliar2 = lDescFranquias(li)
        '    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        '    lCountLine = lCountLine + 1
        'Next li

        '231 a 240. Importante. (10 linhas)
        Print #pNumArqDados, " "
        Print #pNumArqDados, "IMPORTANTE"
        lCountLine = lCountLine + 2
        Call mpTextoImportante(lobjstcPedLoc, pobjstcPed, lDescImportante, "MULTPRO4", 28 - lCountLine)
        For lI = 0 To UBound(lDescImportante) - 1
            If lDescImportante(lI) = "" Then Exit For
            lAuxiliar2 = lDescImportante(lI)
            If gM46V999.gfPreenchido(lAuxiliar2) Then
                lCountLine = lCountLine + 1
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            End If
        Next lI
        Call mpTextoObs(lobjstcPedLoc, lDescImportante, 28 - lCountLine)
        For lI = 0 To UBound(lDescImportante) - 1
            If lDescImportante(lI) = "" Then Exit For
            lAuxiliar2 = lDescImportante(lI)
            If gM46V999.gfPreenchido(lAuxiliar2) Then
                lCountLine = lCountLine + 1
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            End If
        Next lI
        '046 a 051 . Cosseguro
        lAuxiliar2 = ""
        If pobjstcPed.ColecaoPedCoss.Count > 0 And lCountLine < 28 Then
            Print #pNumArqDados, " "
            lCountLine = lCountLine + 1
            For Each lobjstcCoss In pobjstcPed.ColecaoPedCoss
                lAuxiliar2 = lAuxiliar2 & " - ( " & lobjstcCoss.Cod_Coss & " - " & lobjstcCoss.Per_Coss & "% )"
            Next lobjstcCoss
            lCountLine = lCountLine + 1
            Print #pNumArqDados, "COSSEGURO: " & lAuxiliar2
        End If

        If lobjstcPedLoc.ColecaoPedSeguro.Count > 0 And lCountLine < 28 Then
            lAuxiliar2 = ""
            Print #pNumArqDados, " "
            lCountLine = lCountLine + 1
            For Each lobjstcOutros In lobjstcPedLoc.ColecaoPedSeguro
                If Not gM46V999.gfPreenchido(lAuxiliar2) Then
                    lAuxiliar2 = "OUTROS SEGUROS:(" & lobjstcOutros.Cod_Seguradora & "-Ap:" & lobjstcOutros.Num_Apol_Ant & "nº" & lobjstcOutros.Num_Item_Ant & ")"
                Else
                    lAuxiliar2 = lAuxiliar2 & "-(" & lobjstcOutros.Cod_Seguradora & "-Ap:" & lobjstcOutros.Num_Apol_Ant & "nº" & lobjstcOutros.Num_Item_Ant & ")"
                End If
            Next lobjstcOutros
            If gM46V999.gfPreenchido(lAuxiliar2) And lCountLine < 28 Then
                lCountLine = lCountLine + 1
                Print #pNumArqDados, lAuxiliar2
            End If
        End If
        Do While lCountLine < 28
            Print #pNumArqDados, " "
            lCountLine = lCountLine + 1
        Loop
        '095. Local e data.
        lAuxiliar2 = "São Paulo, " & _
                     Format$(Now, "dd") & " de " & Format$(Now, "mmmm") & " de " & Format$(Now, "yyyy")
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '096. Nome do corretor (assinatura)
        lAuxiliar2 = lObjstcCorr.gNom_Corretor
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    Next lobjstcPedLoc
    mfCriDados_MULTPRO4 = True
End Function
Private Function mfCriDados_MULTPROB(ByVal pNumArqDados As Integer, ByVal pobjstcPed As stcA46V702B) As Boolean
'Função: cria arquivo de dados - proposta coletiva - dados do item

'Parâmetro de entrada...pNumArqDados...Número do arquivo de dados.

    Dim lAuxiliar1 As Double  'Auxiliar - utilizado apenas como exemplo.
    Dim lAuxiliar2 As String  'Auxiliar - utilizado apenas como exemplo.
    Dim lAuxiliar3 As String  'Auxiliar - utilizado apenas como exemplo.
    Dim lCodCobertura As Integer    'Código da cobertura.
    Dim lCountLine As Integer    'Contador de linhas.
    Dim lQuant As Integer    'Quantidade de itens.
    Dim lI As Integer    'Utilizado com For...Next.
    Dim lTamanho As Integer    'Tamanho.
    Dim lobjstcBem As stcA46V715B    'Classe (stc) de bens).
    Dim lObjstcCorr As stcA46V705B    'Classe (stc) de corretor.
    Dim lobjstcPedLoc As stcA46V708B    'Classe (stc) de itens.

    Dim lbdTab_Tipo_Fator As ADODB.Connection   'Arquivo: P00Tarifa.
    Dim lrsTab_Tipo_Fator As clsYasRecordSet    'Registro :Tab_Tipo_Fator.
    Dim lobjstcFator As stcA46V720B   'Objeto da classe de estrutura stcA46V720B (TAB_PED_FAT_RISCO).
    Dim lblnAchou As Boolean    'Indica se achou  o fator ou não.
    Dim lMensagem As String
    Dim lSelect As String

    mfCriDados_MULTPROB = False
    lQuant = 0
    For Each lobjstcPedLoc In pobjstcPed.ColecaoPedLoc

        If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTab_Tipo_Fator, lMensagem) = False Then
            Call gM46V999.gpGraLog(2, lMensagem)
            Exit Function
        End If

        'Ficha 1000601
        If gM46V999.gCod_Ramo = e_CodRamo_Residencial Then
            If lobjstcPedLoc.gLiberaDescontoEspecial = True Then
                lSelect = " Select distinct DSC_FATOR, COD_FATOR  from Tbm_Tipo_Fator_Protec "
                lSelect = lSelect & " Where cod_Ramo = " & pobjstcPed.Cod_Ramo
                lSelect = lSelect & " and Cod_Modalidade = " & lobjstcPedLoc.Cod_Produto
                lSelect = lSelect & " and COD_TARIF = " & pobjstcPed.Cod_Produto
                lSelect = lSelect & " and Cod_Sequencia <= " & pobjstcPed.Cod_Sequencia
                lSelect = lSelect & " order by DSC_FATOR"
            Else
                lSelect = " Select distinct DSC_FATOR, COD_FATOR  from Tbm_Tipo_Fator_Protec "
                lSelect = lSelect & " Where cod_Ramo = " & pobjstcPed.Cod_Ramo
                lSelect = lSelect & " and Cod_Modalidade = " & lobjstcPedLoc.Cod_Produto
                lSelect = lSelect & " and COD_TARIF = " & pobjstcPed.Cod_Produto
                lSelect = lSelect & " and Cod_Sequencia <= " & pobjstcPed.Cod_Sequencia
                lSelect = lSelect & " and Cod_Fator <> 7 "    'Não carregar o item Desconto Especial
                lSelect = lSelect & " order by DSC_FATOR"
            End If
        Else
            lSelect = " Select distinct DSC_FATOR, COD_FATOR  from Tbm_Tipo_Fator_Protec "
            lSelect = lSelect & " Where cod_Ramo = " & pobjstcPed.Cod_Ramo
            lSelect = lSelect & " and Cod_Modalidade = " & lobjstcPedLoc.Cod_Produto
            lSelect = lSelect & " and COD_TARIF = " & pobjstcPed.Cod_Produto
            lSelect = lSelect & " and Cod_Sequencia <= " & pobjstcPed.Cod_Sequencia
            lSelect = lSelect & " order by DSC_FATOR"
        End If

        If gM46V999.gfObtRegistro(lbdTab_Tipo_Fator, lSelect, lrsTab_Tipo_Fator, lMensagem) = False Then
            Call gM46V999.gpFecha2(lbdTab_Tipo_Fator)
            Call gM46V999.gpGraLog(2, lMensagem)
            Screen.MousePointer = vbDefault
            Exit Function
        End If

        If lobjstcPedLoc.ColecaoPedLocBem.Count > 0 Or lrsTab_Tipo_Fator.EOF = False Then
            '001 . Título do documento.
            'Print #pNumArqDados, "ANEXO DA PROPOSTA DE SEGURO YASUDA " & UCase(pObjstcPed.gDesc_Ramo)
            Print #pNumArqDados, UCase(pobjstcPed.gDesc_Ramo)

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
            'Print #pNumArqDados, "Emissão: " & Format$(Now, "dd/mm/yyyy hh:nn:ss")
            Print #pNumArqDados, Format$(Now, "dd/mm/yyyy hh:nn:ss")

            lQuant = lQuant + 1
            '004. Número da página.
            'Print #pNumArqDados, "Página " & lQuant
            Print #pNumArqDados, lQuant

            '005. Número do pedido .
            'Print #pNumArqDados, pObjstcPed.Num_Ped & "      Item: " & lObjstcPedLoc.Num_Item
            Print #pNumArqDados, pobjstcPed.Num_Ped

            '006. Número do item .
            Print #pNumArqDados, Format$(lobjstcPedLoc.Num_Item, "0000")

            '007. Corretor (código e nome) .
            If pobjstcPed.ColecaoPedCorr.Count > 0 Then
                Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(1)
                lAuxiliar1 = lObjstcCorr.Cod_Corr
                lAuxiliar2 = lObjstcCorr.gNom_Corretor
            Else
                lAuxiliar1 = 0
                lAuxiliar2 = " "
            End If
            Print #pNumArqDados, Format$(lAuxiliar1, "0000") & " - " & lAuxiliar2

            '008. Tarifa .
            lAuxiliar2 = pobjstcPed.Cod_Produto
            If pobjstcPed.Cod_Produto = 976 And pobjstcPed.Cod_Ramo = e_CodRamo_Empresarial Then
                lAuxiliar2 = lAuxiliar2 & " A"
            End If
            If pobjstcPed.Cod_Sequencia = 1 Then
                lAuxiliar2 = lAuxiliar2 & " (A)"
            End If
            Print #pNumArqDados, lAuxiliar2


            '009. Nosso número .
            Print #pNumArqDados, pobjstcPed.NOSSO_NUMERO
            Print #pNumArqDados, pobjstcPed.Cod_Unicidade & " - " & pobjstcPed.Cod_Ver_Unici

            '010. Local em risco - logradouro .
            lAuxiliar2 = lobjstcPedLoc.Cod_Tip_Logr_Risco & " " & lobjstcPedLoc.Nom_Logr_Risco
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

            '011. Local em risco - número .
            lAuxiliar2 = lobjstcPedLoc.objstcPedInsp.Num_Logr_Risco
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

            '012. Local em risco - complemento .
            lAuxiliar2 = lobjstcPedLoc.Dsc_Compl_Risco
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

            '013. Local em risco - bairro .
            lAuxiliar2 = lobjstcPedLoc.Nom_Bairro_Risco
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

            '014. Local em risco - cidade .
            lAuxiliar2 = lobjstcPedLoc.Nom_Cid_Risco
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

            '015. Local em risco - UF .
            lAuxiliar2 = lobjstcPedLoc.Sig_UF_Risco
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

            '016. Local em risco - CEP .
            lAuxiliar2 = lobjstcPedLoc.gCEP_Formatado
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

            '017 a 041. Fatores de Proteção - linhas de 01 a 25.

            'If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTab_Tipo_Fator, lMensagem) = False Then
            '    Call gM46V999.gpGraLog(2, lMensagem)
            '    Exit Function
            'End If

            'lSelect = " Select distinct DSC_FATOR, COD_FATOR  from Tbm_Tipo_Fator_Protec "
            'lSelect = lSelect & " Where cod_Ramo = " & pObjstcPed.Cod_Ramo
            'lSelect = lSelect & " and Cod_Modalidade = " & lObjstcPedLoc.Cod_Plano
            'lSelect = lSelect & " and COD_TARIF = " & pObjstcPed.Cod_Produto
            'lSelect = lSelect & " and Cod_Sequencia = " & pObjstcPed.Cod_Sequencia
            ''lSelect = lSelect & " and COD_TARIF_INI <= " & pObjstcPed.Cod_Produto
            ''lSelect = lSelect & " and COD_TARIF_TER > " & pObjstcPed.Cod_Produto
            'lSelect = lSelect & " order by DSC_FATOR"

            'If gM46V999.gfObtRegistro(lbdTab_Tipo_Fator, lSelect, lrsTab_Tipo_Fator, lMensagem) = False Then
            '    Call gM46V999.gpFecha2(lbdTab_Tipo_Fator)
            '    Call gM46V999.gpGraLog(2, lMensagem)
            '    Screen.MousePointer = vbDefault
            '    Exit Function
            'End If

            lCountLine = 0
            Do While Not lrsTab_Tipo_Fator.EOF
                lCountLine = lCountLine + 1
                lblnAchou = False

                For Each lobjstcFator In lobjstcPedLoc.ColecaoPedFator
                    If lobjstcFator.Cod_Fator = lrsTab_Tipo_Fator("Cod_Fator") Then
                        lblnAchou = True
                        lAuxiliar2 = Format(lCountLine, "00") & " (X) " & Trim(lrsTab_Tipo_Fator("Dsc_Fator"))
                        If Trim(lobjstcFator.Dsc_Compl_Fator) <> "" Then
                            lAuxiliar2 = lAuxiliar2 & " - " & Trim(lobjstcFator.Dsc_Compl_Fator)
                        End If
                        Print #pNumArqDados, lAuxiliar2
                        Exit For
                    End If
                Next lobjstcFator

                If lblnAchou = False Then
                    lAuxiliar2 = Format(lCountLine, "00") & " ( ) " & Trim(lrsTab_Tipo_Fator("Dsc_Fator"))
                    Print #pNumArqDados, lAuxiliar2
                End If
                lrsTab_Tipo_Fator.MoveNext
            Loop

            Call gM46V999.gpFecha2(lbdTab_Tipo_Fator)

            Do While lCountLine < 25
                lCountLine = lCountLine + 1
                'Print #pNumArqDados, lCountLine & " () testesssss"
                Print #pNumArqDados, " "
            Loop

            '042 a 086. Fatores Bens - linhas de 01 a 45.
            '         Print #pNumArqDados, "Bens:"
            lCountLine = 0
            For Each lobjstcBem In lobjstcPedLoc.ColecaoPedLocBem
                lCountLine = lCountLine + 1
                'lAuxiliar2 = Format(lCountLine, "000") & " - " & Left(lObjstcBem.Dsc_Bem, 100) & String(100 - Len(lObjstcBem.Dsc_Bem), ".") & "R$ " & Format(lObjstcBem.Val_Bem, "standard")
                lAuxiliar3 = lobjstcBem.gTipBem & " - " & lobjstcBem.Dsc_Bem
                lAuxiliar2 = Format(lCountLine, "000") & " - " & Left(lAuxiliar3, 100) & String(90 - Len(Left(lAuxiliar3, 90)), ".") & "R$ " & Format(lobjstcBem.Val_Bem, "standard")
                Print #pNumArqDados, lAuxiliar2
            Next lobjstcBem

            'Do While lCountLine < 75
            Do While lCountLine < 45
                'Print #pNumArqDados, " "

                lCountLine = lCountLine + 1
                'Print #pNumArqDados, lCountLine & "BENS BENS BENS"
                Print #pNumArqDados, " "
            Loop

            '095. Local e data.
            lAuxiliar2 = "São Paulo, " & _
                         Format$(Now, "dd") & " de " & Format$(Now, "mmmm") & " de " & Format$(Now, "yyyy")
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

            '096. Nome do corretor (assinatura)
            lAuxiliar2 = lObjstcCorr.gNom_Corretor
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        End If
    Next lobjstcPedLoc

    If lQuant > 0 Then
        mfCriDados_MULTPROB = True
    Else
        Call gM46V999.gpFecha2(lbdTab_Tipo_Fator)
    End If

End Function
Private Function mfCriDados_MULTPROC(ByVal pNumArqDados As Integer, ByVal pobjstcPed As stcA46V702B) As Boolean
'Função: cria arquivo de dados - proposta coletiva - dados do item

'Parâmetro de entrada...pNumArqDados...Número do arquivo de dados.

    Dim lAuxiliar1 As Double  'Auxiliar - utilizado apenas como exemplo.
    Dim lAuxiliar2 As String  'Auxiliar - utilizado apenas como exemplo.
    Dim lAuxiliar3 As String  'Auxiliar - utilizado apenas como exemplo.
    Dim lCodCobertura As Integer    'Código da cobertura.
    Dim lCountLine As Integer    'Contador de linhas.
    Dim lQuant As Integer    'Quantidade de itens.
    Dim lI As Integer    'Utilizado com For...Next.
    Dim lTamanho As Integer    'Tamanho.
    Dim lObjstcCorr As stcA46V705B    'Classe (stc) de corretor.
    Dim lobjstcPedLoc As stcA46V708B    'Classe (stc) de itens.
    Dim lobjstcCobert As stcA46V704B
    Dim lobjstcTextoEspec As stcA46V726B    'Objeto da classe de estrutura stcA46V726B (TAB_PED_ESPEC).
    Dim lblnAchou As Boolean    'Indica se achou  o fator ou não.
    Dim lMensagem As String
    Dim lTemInfo As Boolean
    Dim lbdTbm_Cob_LucroCess As ADODB.Connection    'Arquivo: P00Multi.
    Dim lrsTbm_Cob_LucroCess As clsYasRecordSet     'Registro : Tbm_Cob_LucroCess.
    Dim lSelect As String

    mfCriDados_MULTPROC = False
    lQuant = 0
    If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTbm_Cob_LucroCess, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Function
    End If
    lTemInfo = False
    For Each lobjstcPedLoc In pobjstcPed.ColecaoPedLoc
        '001 . Título do documento.
        'Print #pNumArqDados, "ANEXO DA PROPOSTA DE SEGURO YASUDA " & UCase(pObjstcPed.gDesc_Ramo)
        Print #pNumArqDados, UCase(pobjstcPed.gDesc_Ramo)

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
        'Print #pNumArqDados, "Emissão: " & Format$(Now, "dd/mm/yyyy hh:nn:ss")
        Print #pNumArqDados, Format$(Now, "dd/mm/yyyy hh:nn:ss")

        lQuant = lQuant + 1
        '004. Número da página.
        'Print #pNumArqDados, "Página " & lQuant
        Print #pNumArqDados, lQuant

        '005. Número do pedido .
        'Print #pNumArqDados, pObjstcPed.Num_Ped & "      Item: " & lObjstcPedLoc.Num_Item
        Print #pNumArqDados, pobjstcPed.Num_Ped

        '006. Número do item .
        Print #pNumArqDados, Format$(lobjstcPedLoc.Num_Item, "0000")

        '007. Corretor (código e nome) .
        If pobjstcPed.ColecaoPedCorr.Count > 0 Then
            Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(1)
            lAuxiliar1 = lObjstcCorr.Cod_Corr
            lAuxiliar2 = lObjstcCorr.gNom_Corretor
        Else
            lAuxiliar1 = 0
            lAuxiliar2 = " "
        End If
        Print #pNumArqDados, Format$(lAuxiliar1, "0000") & " - " & lAuxiliar2

        '008. Tarifa .
        lAuxiliar2 = pobjstcPed.Cod_Produto
        If pobjstcPed.Cod_Produto = 976 And pobjstcPed.Cod_Ramo = e_CodRamo_Empresarial Then
            lAuxiliar2 = lAuxiliar2 & " A"
        End If
        If pobjstcPed.Cod_Sequencia = 1 Then
            lAuxiliar2 = lAuxiliar2 & " (A)"
        End If
        Print #pNumArqDados, lAuxiliar2

        '009. Nosso número .
        Print #pNumArqDados, pobjstcPed.NOSSO_NUMERO
        Print #pNumArqDados, pobjstcPed.Cod_Unicidade & " - " & pobjstcPed.Cod_Ver_Unici

        '010. Local em risco - logradouro .
        lAuxiliar2 = lobjstcPedLoc.Cod_Tip_Logr_Risco & " " & lobjstcPedLoc.Nom_Logr_Risco
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '011. Local em risco - número .
        lAuxiliar2 = lobjstcPedLoc.objstcPedInsp.Num_Logr_Risco
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '012. Local em risco - complemento .
        lAuxiliar2 = lobjstcPedLoc.Dsc_Compl_Risco
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '013. Local em risco - bairro .
        lAuxiliar2 = lobjstcPedLoc.Nom_Bairro_Risco
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '014. Local em risco - cidade .
        lAuxiliar2 = lobjstcPedLoc.Nom_Cid_Risco
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '015. Local em risco - UF .
        lAuxiliar2 = lobjstcPedLoc.Sig_UF_Risco
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '016. Local em risco - CEP .
        lAuxiliar2 = lobjstcPedLoc.gCEP_Formatado
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '017 a 036. Beneficiarios - linhas de 01 a 20.
        lCountLine = 0
        If pobjstcPed.ColecaoTextoEspec.Count > 0 Then
            For Each lobjstcTextoEspec In pobjstcPed.ColecaoTextoEspec

                If InStr(1, lobjstcTextoEspec.Espec_lin1, "** FRANQUIA **", vbTextCompare) > 0 Or _
                   InStr(1, lobjstcTextoEspec.Espec_lin2, "** FRANQUIA **", vbTextCompare) > 0 Then
                    Do While lCountLine < 12
                        Print #pNumArqDados, " "
                        lCountLine = lCountLine + 1
                    Loop
                    Exit For
                End If

                Print #pNumArqDados, lobjstcTextoEspec.Espec_lin1
                Print #pNumArqDados, IIf(lobjstcTextoEspec.Espec_lin2 = "", " ", lobjstcTextoEspec.Espec_lin2)
                Print #pNumArqDados, IIf(lobjstcTextoEspec.Espec_lin3 = "", " ", lobjstcTextoEspec.Espec_lin3)
                Print #pNumArqDados, IIf(lobjstcTextoEspec.Espec_lin4 = "", " ", lobjstcTextoEspec.Espec_lin4)
                Print #pNumArqDados, IIf(lobjstcTextoEspec.Espec_lin5 = "", " ", lobjstcTextoEspec.Espec_lin5)
                Print #pNumArqDados, IIf(lobjstcTextoEspec.Espec_lin6 = "", " ", lobjstcTextoEspec.Espec_lin6)
                Print #pNumArqDados, IIf(lobjstcTextoEspec.Espec_lin7 = "", " ", lobjstcTextoEspec.Espec_lin7)
                Print #pNumArqDados, IIf(lobjstcTextoEspec.Espec_lin8 = "", " ", lobjstcTextoEspec.Espec_lin8)
                Print #pNumArqDados, IIf(lobjstcTextoEspec.Espec_lin9 = "", " ", lobjstcTextoEspec.Espec_lin9)
                Print #pNumArqDados, IIf(lobjstcTextoEspec.Espec_lin10 = "", " ", lobjstcTextoEspec.Espec_lin10)
                Print #pNumArqDados, IIf(lobjstcTextoEspec.Espec_lin11 = "", " ", lobjstcTextoEspec.Espec_lin11)
                Print #pNumArqDados, IIf(lobjstcTextoEspec.Espec_lin12 = "", " ", lobjstcTextoEspec.Espec_lin12)
                lTemInfo = True
            Next lobjstcTextoEspec
            lCountLine = 12
        End If


        Do While lCountLine < 20
            lCountLine = lCountLine + 1
            Print #pNumArqDados, " "
        Loop

        '042 a 086. Lucros Cessantes - linhas de 01 a 08.

        lCountLine = 0
        For Each lobjstcCobert In lobjstcPedLoc.ColecaoPedCobert
            If lobjstcCobert.Num_Item = lobjstcPedLoc.Num_Item Then
                '1 - Procura pela Cobertura na Tabela de Coeficiente
                'Se encontrar é deve imprimir

                lSelect = " Select *  from Tbm_Cob_LucroCess "
                lSelect = lSelect & " Where Cod_Ramo = " & lobjstcCobert.Cod_Ramo
                lSelect = lSelect & " and Cod_Cobert = " & lobjstcCobert.Cod_Cobert
                lSelect = lSelect & " and Cod_Tarif = " & pobjstcPed.Cod_Produto
                lSelect = lSelect & " and Cod_Sequencia <= " & pobjstcPed.Cod_Sequencia

                If gM46V999.gfObtRegistro(lbdTbm_Cob_LucroCess, lSelect, lrsTbm_Cob_LucroCess, lMensagem) = False Then
                    Call gM46V999.gpFecha2(lbdTbm_Cob_LucroCess)
                    Call gM46V999.gpGraLog(2, lMensagem)
                    Screen.MousePointer = vbDefault
                    Exit Function
                End If

                '2 - Se encontrar permite cadastro do Valor de Risco de Lucro Cessantes
                If Not lrsTbm_Cob_LucroCess.EOF Then
                    lCountLine = lCountLine + 1
                    lAuxiliar3 = "Cobertura: " & lobjstcCobert.gNomeCobertura & " - Valor em Risco: "
                    lAuxiliar2 = Left(lAuxiliar3, 110) & String(100 - Len(lAuxiliar3), ".") & "R$ " & Format(lobjstcCobert.Val_Risco_Cob, "standard")
                    Print #pNumArqDados, lAuxiliar2
                    lTemInfo = True
                End If
            End If
        Next lobjstcCobert


        Do While lCountLine < 8
            lCountLine = lCountLine + 1
            Print #pNumArqDados, " "
        Loop

        '095. Local e data.
        lAuxiliar2 = "São Paulo, " & _
                     Format$(Now, "dd") & " de " & Format$(Now, "mmmm") & " de " & Format$(Now, "yyyy")
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '096. Nome do corretor (assinatura)
        lAuxiliar2 = lObjstcCorr.gNom_Corretor
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    Next lobjstcPedLoc

    If lTemInfo = True Then
        mfCriDados_MULTPROC = True
    End If
    Call gM46V999.gpFecha2(lbdTbm_Cob_LucroCess)


End Function

Private Function mfCriDados_MULTPROD(ByVal pNumArqDados As Integer, ByVal pobjstcPed As stcA46V702B) As Boolean
'Função: cria arquivo de dados - proposta coletiva - dados do item

'Parâmetro de entrada...pNumArqDados...Número do arquivo de dados.

    Dim lAuxiliar1 As Double  'Auxiliar - utilizado apenas como exemplo.
    Dim lAuxiliar2 As String  'Auxiliar - utilizado apenas como exemplo.
    Dim lAuxiliar3 As String  'Auxiliar - utilizado apenas como exemplo.
    Dim lCodCobertura As Integer    'Código da cobertura.
    Dim lCountLine As Integer    'Contador de linhas.
    Dim lQuant As Integer    'Quantidade de itens.
    Dim lI As Integer    'Utilizado com For...Next.
    Dim lTamanho As Integer    'Tamanho.
    Dim lObjstcCorr As stcA46V705B    'Classe (stc) de corretor.
    Dim lobjstcPedLoc As stcA46V708B    'Classe (stc) de itens.
    Dim lobjstcCobert As stcA46V704B
    Dim lobjstcTextoEspec As stcA46V726B    'Objeto da classe de estrutura stcA46V726B (TAB_PED_ESPEC).
    Dim lobjstcBem As stcA46V715B    'Classe (stc) de bens).
    Dim lblnAchou As Boolean    'Indica se achou  o fator ou não.
    Dim lMensagem As String
    Dim lTemInformacao As Boolean
    Dim lSelect As String
    Dim lbdTbm_Cob_LucroCess As ADODB.Connection    'Arquivo: P00Multi.
    Dim lrsTbm_Cob_LucroCess As clsYasRecordSet     'Registro : Tbm_Cob_LucroCess.


    mfCriDados_MULTPROD = False
    lQuant = 0
    If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTbm_Cob_LucroCess, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Function
    End If
    lTemInformacao = False
    For Each lobjstcPedLoc In pobjstcPed.ColecaoPedLoc
        '001 . Título do documento.
        'Print #pNumArqDados, "ANEXO DA PROPOSTA DE SEGURO YASUDA " & UCase(pObjstcPed.gDesc_Ramo)
        Print #pNumArqDados, UCase(pobjstcPed.gDesc_Ramo)

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
        'Print #pNumArqDados, "Emissão: " & Format$(Now, "dd/mm/yyyy hh:nn:ss")
        Print #pNumArqDados, Format$(Now, "dd/mm/yyyy hh:nn:ss")

        lQuant = lQuant + 1
        '004. Número da página.
        'Print #pNumArqDados, "Página " & lQuant

        '005. Número do pedido .
        'Print #pNumArqDados, pObjstcPed.Num_Ped & "      Item: " & lObjstcPedLoc.Num_Item
        Print #pNumArqDados, pobjstcPed.Num_Ped

        '006. Número do item .
        Print #pNumArqDados, Format$(lobjstcPedLoc.Num_Item, "0000")

        '007. Corretor (código e nome) .
        If pobjstcPed.ColecaoPedCorr.Count > 0 Then
            Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(1)
            lAuxiliar1 = lObjstcCorr.Cod_Corr
            lAuxiliar2 = lObjstcCorr.gNom_Corretor
        Else
            lAuxiliar1 = 0
            lAuxiliar2 = " "
        End If
        Print #pNumArqDados, Format$(lAuxiliar1, "0000") & " - " & lAuxiliar2

        '008. Tarifa .
        lAuxiliar2 = pobjstcPed.Cod_Produto
        If pobjstcPed.Cod_Produto = 976 And pobjstcPed.Cod_Ramo = e_CodRamo_Empresarial Then
            lAuxiliar2 = lAuxiliar2 & " A"
        End If
        If pobjstcPed.Cod_Sequencia = 1 Then
            lAuxiliar2 = lAuxiliar2 & " (A)"
        End If
        Print #pNumArqDados, lAuxiliar2

        '009. Nosso número .
        Print #pNumArqDados, pobjstcPed.NOSSO_NUMERO
        Print #pNumArqDados, pobjstcPed.Cod_Unicidade & " - " & pobjstcPed.Cod_Ver_Unici

        lCountLine = 0
        For Each lobjstcBem In lobjstcPedLoc.ColecaoPedLocBem
            lTemInformacao = True
            lCountLine = lCountLine + 1
            lAuxiliar3 = lobjstcBem.gTipBem & " - " & lobjstcBem.Dsc_Bem
            lAuxiliar2 = Format(lCountLine, "000") & " - " & Left(lAuxiliar3, 100) & String(90 - Len(Left(lAuxiliar3, 90)), ".") & "R$ " & Format(lobjstcBem.Val_Bem, "standard")
            Print #pNumArqDados, lAuxiliar2
        Next lobjstcBem

        Do While lCountLine < 49
            lCountLine = lCountLine + 1
            Print #pNumArqDados, " "
        Loop

        '017 a 036. Beneficiarios - linhas de 01 a 20.
        lCountLine = 0
        If pobjstcPed.ColecaoTextoEspec.Count > 0 Then
            lTemInformacao = True
            For Each lobjstcTextoEspec In pobjstcPed.ColecaoTextoEspec

                If InStr(1, lobjstcTextoEspec.Espec_lin1, "** FRANQUIA **", vbTextCompare) > 0 Or _
                   InStr(1, lobjstcTextoEspec.Espec_lin2, "** FRANQUIA **", vbTextCompare) > 0 Then
                    Exit For
                End If

                Print #pNumArqDados, lobjstcTextoEspec.Espec_lin1 & " "
                Print #pNumArqDados, lobjstcTextoEspec.Espec_lin2 & " "
                Print #pNumArqDados, lobjstcTextoEspec.Espec_lin3 & " "
                Print #pNumArqDados, lobjstcTextoEspec.Espec_lin4 & " "
                Print #pNumArqDados, lobjstcTextoEspec.Espec_lin5 & " "
                Print #pNumArqDados, lobjstcTextoEspec.Espec_lin6 & " "
                Print #pNumArqDados, lobjstcTextoEspec.Espec_lin7 & " "
                Print #pNumArqDados, lobjstcTextoEspec.Espec_lin8 & " "
                Print #pNumArqDados, lobjstcTextoEspec.Espec_lin9 & " "
                Print #pNumArqDados, lobjstcTextoEspec.Espec_lin10 & " "
                Print #pNumArqDados, lobjstcTextoEspec.Espec_lin11 & " "
                Print #pNumArqDados, lobjstcTextoEspec.Espec_lin12 & " "
                lCountLine = lCountLine + 12
            Next lobjstcTextoEspec
        End If

        Do While lCountLine < 20
            Print #pNumArqDados, " "
            lCountLine = lCountLine + 1
        Loop

        '042 a 086. Lucros Cessantes
        lCountLine = 0
        For Each lobjstcCobert In lobjstcPedLoc.ColecaoPedCobert
            If lobjstcCobert.Num_Item = lobjstcPedLoc.Num_Item Then
                '1 - Procura pela Cobertura na Tabela de Coeficiente
                'Se encontrar é deve imprimir

                lSelect = " Select *  from Tbm_Cob_LucroCess "
                lSelect = lSelect & " Where Cod_Ramo = " & lobjstcCobert.Cod_Ramo
                lSelect = lSelect & " and Cod_Cobert = " & lobjstcCobert.Cod_Cobert
                lSelect = lSelect & " and Cod_Tarif = " & pobjstcPed.Cod_Produto
                lSelect = lSelect & " and Cod_Sequencia <= " & pobjstcPed.Cod_Sequencia

                If gM46V999.gfObtRegistro(lbdTbm_Cob_LucroCess, lSelect, lrsTbm_Cob_LucroCess, lMensagem) = False Then
                    Call gM46V999.gpFecha2(lbdTbm_Cob_LucroCess)
                    Call gM46V999.gpGraLog(2, lMensagem)
                    Screen.MousePointer = vbDefault
                    Exit Function
                End If

                '2 - Se encontrar permite cadastro do Valor de Risco de Lucro Cessantes
                If Not lrsTbm_Cob_LucroCess.EOF Then
                    lTemInformacao = True
                    lCountLine = lCountLine + 1
                    lAuxiliar3 = "Cobertura: " & lobjstcCobert.gNomeCobertura & " - Valor em Risco: "
                    lAuxiliar2 = Left(lAuxiliar3, 110) & String(100 - Len(lAuxiliar3), ".") & "R$ " & Format(lobjstcCobert.Val_Risco_Cob, "standard")
                    Print #pNumArqDados, lAuxiliar2
                End If
            End If
        Next lobjstcCobert

        Do While lCountLine < 8
            lCountLine = lCountLine + 1
            Print #pNumArqDados, " "
        Loop

        '095. Local e data.
        lAuxiliar2 = "São Paulo, " & _
                     Format$(Now, "dd") & " de " & Format$(Now, "mmmm") & " de " & Format$(Now, "yyyy")
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '096. Nome do corretor (assinatura)
        lAuxiliar2 = lObjstcCorr.gNom_Corretor
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    Next lobjstcPedLoc

    If lQuant > 0 And lTemInformacao = True Then
        mfCriDados_MULTPROD = True
    End If
    Call gM46V999.gpFecha2(lbdTbm_Cob_LucroCess)

End Function





Private Sub mpTextoObs(pobjstcPedLoc As stcA46V708B, pTexto() As String, pTamArray As Integer)
    Dim lI As Integer
    Dim lObs As Variant
    Dim lCount As Integer
    Dim lAuxiliar As String
    Dim lPos As Long
    Dim lObservacao As String
    Dim lobjstcClau As stcA46V703B


    ReDim pTexto(pTamArray)
    lI = 0
    If gM46V999.gfPreenchido(pobjstcPedLoc.objstcPedObs.Observacao) Then
        pobjstcPedLoc.objstcPedObs.Observacao = pobjstcPedLoc.objstcPedObs.Observacao & "&%"
    End If
    lObservacao = pobjstcPedLoc.objstcPedObs.Observacao

    For Each lobjstcClau In pobjstcPedLoc.ColecaoPedClau
        'Ficha 1001013 - Imprimir Todas as Clausulas
        '        If lobjstcClau.Cod_Clau >= 310 Or (lobjstcClau.Cod_Ramo = e_CodRamo_Residencial And lobjstcClau.Cod_Clau > 300) Then
        lObservacao = lObservacao & "- Cláusula " & lobjstcClau.Cod_Clau & " - " & lobjstcClau.gDescClausula & "&%"
        '        End If
    Next lobjstcClau

    lObs = Split(lObservacao, "&%")
    For lCount = 0 To UBound(lObs)
        lAuxiliar = lObs(lCount)
        If gM46V999.gfPreenchido(lAuxiliar) Then
            If Len(lAuxiliar) < 128 Then
                pTexto(lI) = lAuxiliar
                lI = lI + 1
                If lI > pTamArray Then Exit For
            Else
                lPos = 1
                Do
                    pTexto(lI) = Mid(lAuxiliar, lPos, 128)
                    lI = lI + 1
                    If lI > pTamArray Then Exit For
                    If Len(Mid(lAuxiliar, lPos, 128)) < 128 Then Exit Do
                    lPos = lPos + 128
                Loop
            End If
        End If
    Next lCount

End Sub
Private Sub mpTextoObsFonte12(pobjstcPedLoc As stcA46V708B, pTexto() As String, pTamArray As Integer)
    Dim lI As Integer
    Dim lObs As Variant
    Dim lCount As Integer
    Dim lAuxiliar As String
    Dim lPos As Long
    Dim lObservacao As String
    Dim lobjstcClau As stcA46V703B


    ReDim pTexto(pTamArray)
    lI = 0
    If gM46V999.gfPreenchido(pobjstcPedLoc.objstcPedObs.Observacao) Then
        pobjstcPedLoc.objstcPedObs.Observacao = pobjstcPedLoc.objstcPedObs.Observacao & "&%"
    End If
    lObservacao = pobjstcPedLoc.objstcPedObs.Observacao

    If gM46V999.gCodProduto < 1014 Then  '1100589 - Tarifa 1014 e 1015 Julho 2011 - foi criado subrotina (mpImpClausulas) para impressao das clausulas
        For Each lobjstcClau In pobjstcPedLoc.ColecaoPedClau
            'Ficha 1001013 - Imprimir Todas as Clausulas
            '        If lobjstcClau.Cod_Clau >= 310 Or (lobjstcClau.Cod_Ramo = e_CodRamo_Residencial And lobjstcClau.Cod_Clau > 300) Then

            lObservacao = lObservacao & "- Cláusula " & lobjstcClau.Cod_Clau & " - " & lobjstcClau.gDescClausula & "&%"
            '        End If
        Next lobjstcClau
    End If

    lObs = Split(lObservacao, "&%")
    For lCount = 0 To UBound(lObs)
        lAuxiliar = lObs(lCount)
        If gM46V999.gfPreenchido(lAuxiliar) Then
            If Len(lAuxiliar) < 100 Then
                pTexto(lI) = lAuxiliar
                lI = lI + 1
                If lI > pTamArray Then Exit For
            Else
                lPos = 1
                Do
                    pTexto(lI) = Mid(lAuxiliar, lPos, 100)
                    lI = lI + 1
                    If lI > pTamArray Then Exit For
                    If Len(Mid(lAuxiliar, lPos, 100)) < 100 Then Exit Do
                    lPos = lPos + 100
                Loop
            End If
        End If
    Next lCount

End Sub

Private Sub mpTextoImportante(pobjstcPedLoc As stcA46V708B, pobjstcPed As stcA46V702B, pTexto() As String, pForm As String, pTamArray As Integer)
    Dim lAuxiliar As String
    Dim lobjstcCobert As stcA46V704B
    Dim lobjstcOutros As stcA46V716B
    Dim lI As Integer
    Dim lCod_Ramo As e_CodRamo
    Dim lObjstcCorr As stcA46V705B
    Dim lTemAssistencia24hr As Boolean
    Dim lobjstcClau As stcA46V703B

    ReDim pTexto(pTamArray)

    lCod_Ramo = pobjstcPedLoc.Cod_Ramo
    If pobjstcPed.Cod_Tip_Doc = e_TipDoc_Cotacao Then
        pTexto(lI) = "Cotação e LMI para simples consulta, não implica na aceitação automática."
        lI = lI + 1
    End If
    '1101045 - Tarifa 1022 e 1023 Novembro 2011
    '''    If pobjstcPed.Cod_Tip_Doc = e_TipDoc_Cotacao Then
    lTemAssistencia24hr = False
    For Each lobjstcClau In pobjstcPedLoc.ColecaoPedClau
        If lobjstcClau.Cod_Clau = 400 Or lobjstcClau.Cod_Clau = 401 Or lobjstcClau.Cod_Clau = 402 Then
            lTemAssistencia24hr = True
        End If
    Next lobjstcClau
    '1101045 - Tarifa 1022 e 1023 Novembro 2011
    If lTemAssistencia24hr = False Then
        '            If mStcA46V708B.ObjPai.Cod_Tip_Doc = e_TipDoc_Cotacao Then
        If lCod_Ramo = e_CodRamo_Residencial Then
            If (val(pobjstcPedLoc.Cod_Clasf) = gM46V999.gfBuscarDePara(1, 2)) And _
               (val(pobjstcPedLoc.Cod_Compl_Clasf) = gM46V999.gfBuscarDePara(1, 2, 13) Or val(pobjstcPedLoc.Cod_Compl_Clasf) = gM46V999.gfBuscarDePara(1, 2, 23)) Then
                'nada
            Else
                pTexto(lI) = "Assistência Domiciliar 24 Horas Gratuita (Exceto para Residências Eventuais)."
                lI = lI + 1
            End If
        ElseIf lCod_Ramo = e_CodRamo_Empresarial And gM46V999.gCodProduto >= 950 Then
            pTexto(lI) = "Assistência 24 Horas Gratuita."
            lI = lI + 1
        ElseIf lCod_Ramo <> e_CodRamo_Roubos Then
            If pobjstcPed.Cod_Tip_Doc = e_TipDoc_Cotacao Then
                pTexto(lI) = "Assistência 24 Horas Gratuita, restrita para apólice individual."
                lI = lI + 1
            End If
        End If
    End If
    '''    End If
    If lCod_Ramo = e_CodRamo_Roubos Then
        pTexto(lI) = "Nos termos do art. 782 do Código Civil Brasileiro, o Segurado obriga-se a declarar à Yasuda Seguros, no momento da apresentação da Proposta "
        lI = lI + 1
        pTexto(lI) = "de Seguro, a existência de quaisquer outros seguros contratados nessa ou em outra seguradora que garantam, contra os mesmos riscos, os bens"
        lI = lI + 1
        pTexto(lI) = "e interesses segurados. O Segurado deve também comunicar à Yasuda Seguros a efetivação posterior de outros seguros conforme acima descrito."
        lI = lI + 1
        pTexto(lI) = "A presente cotação é apresentada apenas para análise de parâmetro de custo não implicando no reconhecimento da aceitação do seguro a qual "
        lI = lI + 1
        pTexto(lI) = "ficará condicionada a verificação das condições favoráveis do(s) risco(s) a ser(em) constatadas mediante inspeção(ões) no(s) local(is). "
        lI = lI + 1
        pTexto(lI) = "A Yasuda Seguros se reserva ainda ao direito de alterar a presente cotação caso o(s) risco(s) em questão não apresente(m) condições favoráveis "
        lI = lI + 1
        pTexto(lI) = "e/ou sejam constatadas omissões/divergências nas informações recebidas."
        lI = lI + 1
        pTexto(lI) = "Cotação válida por 15 dias."
        lI = lI + 1

        If pobjstcPedLoc.ColecaoPedLocBem.Count > 0 Then
            pTexto(lI) = "A soma dos valores dos Bens não poderá ser superior ao valor do LMG Total."
            lI = lI + 1
        End If
        If pobjstcPed.ColecaoPedCorr.Count > 0 Then
            Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(1)
            If lObjstcCorr.Per_Corr = 0 Then
                pTexto(lI) = "Prêmio apresentado na forma NET, ou seja, sem carregamento de Comissão de Corretagem."
            Else
                pTexto(lI) = "Prêmio apresentado na forma Líquida, ou seja, com carregamento de Comissão de Corretagem."
            End If
            lI = lI + 1
        End If
    End If
'        Tarifa 11.12
'    If pobjstcPed.Cod_Tip_Doc = e_TipDoc_Cotacao Then
'        pTexto(lI) = "Na Razão Social não deverá constar a expressão PCPT, conforme informativo Yasuda  01/2004 de 13/04/2004."
'        lI = lI + 1
'    End If
    If pobjstcPed.objstcCtrlEmis.Tip_Cotac = "2" And gM46V999.gfPreenchido(pobjstcPed.objstcCtrlEmis.Num_Cotac) Then
        pTexto(lI) = "SISCOTA:" & pobjstcPed.objstcCtrlEmis.Num_Cotac
        lI = lI + 1
    End If

    '### INSPECAO
    If pobjstcPed.Cod_Tip_Doc = e_TipDoc_Proposta Then
        If pForm <> "MULTPRO2" Then
            pTexto(lI) = "A aceitação desta proposta está condicionada a análise da Yasuda Seguros até 15 dias de seu recebimento. Este prazo ficará "
            lI = lI + 1
            pTexto(lI) = "suspenso até a entrega de documentos complementares e/ou data da inspeção do risco."
            lI = lI + 1
        End If
    Else
        If pobjstcPedLoc.Insp_Obr = True Or pobjstcPedLoc.Ind_Sin_Inspec > 0 Then
            pTexto(lI) = "A aceitação deste seguro está condicionada a solicitação e análise da inspeção do risco, antes da entrada da proposta na cia."
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
                                        pobjstcPedLoc.Cod_Plano = 2102 Or _
                                        pobjstcPedLoc.Cod_Plano = 2142) Then    'Ficha 1000495 - Plano Belacity
        pTexto(lI) = "O presente seguro destina-se a cobrir exclusivamente o prédio (imóvel)."
        lI = lI + 1
    End If



    'Ficha 1001070
    If pobjstcPedLoc.Cod_Plano = 2157 Or pobjstcPedLoc.Cod_Plano = 2168 Or pobjstcPedLoc.Cod_Plano = 2169 Then
        pTexto(lI) = "Cláusula Beneficiária a favor do proprietário do imóvel aplicável exclusivamente ao prédio."
        lI = lI + 1
    End If
    If pobjstcPedLoc.Cod_Plano = 2168 Or pobjstcPedLoc.Cod_Plano = 2169 Then
        pTexto(lI) = "Seguro exclusivamente para o prédio."
        lI = lI + 1
        For Each lobjstcCobert In pobjstcPedLoc.ColecaoPedCobert
            If lobjstcCobert.Cod_Cobert = gM46V999.gfBuscarDePara(5, 1) Then
                pTexto(lI) = "A cobertura 05 Roubo ou Furto Com Vestígios garante exclusivamente os danos materiais diretamente causados ao prédio "
                lI = lI + 1
                pTexto(lI) = "segurado durante a prática ou tentativa de roubo e/ou furto com vestígios."
                lI = lI + 1
            End If
            If lobjstcCobert.Cod_Cobert = gM46V999.gfBuscarDePara(23, 1) Then
                pTexto(lI) = "A cobertura 23 Roubo de Bens garante exclusivamente os danos materiais diretamente causados ao prédio segurado "
                lI = lI + 1
                pTexto(lI) = "durante a prática de roubo ou furto com vestígios, quer o evento se tenha consumado, ou quer se tenha caracterizado"
                lI = lI + 1
                pTexto(lI) = "a simples tentativa."
                lI = lI + 1
            End If
        Next lobjstcCobert
    End If


    'Ficha 901378
    If pForm <> "MULTPRO2" Then
        pTexto(lI) = "O registro deste plano na SUSEP não implica, por parte da Autarquia, incentivo ou recomendação a sua comercialização."
        lI = lI + 1
        pTexto(lI) = "O segurado poderá consultar a situação cadastral de seu corretor de seguros, no site www.susep.gov.br, por meio de seu "
        lI = lI + 1
        pTexto(lI) = "registro na SUSEP, nome completo, CNPJ ou CPF."
        lI = lI + 1
    End If

    'Retirado na nova versão

    'For Each lObjstcCobert In pobjstcPedLoc.ColecaoPedCobert
    '    Select Case lObjstcCobert.Cod_Cobert
    '        Case 39, 40, 41, 42, 43, 44, 47, 50
    '            pTexto(li) = "Período Indenitário: 12 meses. Para prazos inferiores, solicitar cotação à Unidade de produto"
    '            li = li + 1
    '            Exit For
    '    End Select
    'Next lObjstcCobert

    'Ficha 901346
    If val(pobjstcPedLoc.Cod_Clasf) = gM46V999.gfBuscarDePara(22, 2) And val(pobjstcPedLoc.Cod_Compl_Clasf) = gM46V999.gfBuscarDePara(22, 2, 13) Then
        For Each lobjstcCobert In pobjstcPedLoc.ColecaoPedCobert
            '''            Select Case lobjstcCobert.Cod_Cobert
            '''                Case 56
            '''                    pTexto(li) = "Cobertura 56 - Pátios - até 100 Km: Nos sinistros ocorridos com veículos embarcados em guinchos/cegonhas de terceiros,"
            '''                    li = li + 1
            '''                    pTexto(li) = "a Franquia fica substituída pela participação obrigatória do segurado em 10% prej. mín."
            '''                    li = li + 1
            '''                    pTexto(li) = "R$ 1.500,00/R$ 3.000,00 - Nacionais/Importados."
            '''                    li = li + 1
            '''                Case 57
            '''                    pTexto(li) = "Cobertura 57 - R C Concessionárias - até 100 Km: Nos sinistros ocorridos com veículos embarcados em guinchos/cegonhas"
            '''                    li = li + 1
            '''                    pTexto(li) = "de terceiros, a Franquia fica substituída pela participação obrigatória do segurado em 10% prej. mín."
            '''                    li = li + 1
            '''                    pTexto(li) = "R$ 1.500,00/R$ 3.000,00 - Nacionais/Importados."
            '''                    li = li + 1
            '''                Case 122
            '''                    pTexto(li) = "Cobertura 122 - Pátios - até 200 Km: Nos sinistros ocorridos com veículos embarcados em guinchos/cegonhas de terceiros, a Franquia"
            '''                    li = li + 1
            '''                    pTexto(li) = "fica substituída pela participação obrigatória do segurado em 10% prej. mín. R$ 1.500,00/R$ 3.000,00 - Nacionais/Importados."
            '''                    li = li + 1
            '''                Case 123
            '''                    pTexto(li) = "Cobertura 123 - Pátios - até 300 Km : Nos sinistros ocorridos com veículos embarcados em guinchos/cegonhas de terceiros, a Franquia"
            '''                    li = li + 1
            '''                    pTexto(li) = "fica substituída pela participação obrigatória do segurado em 10% prej. mín. R$ 1.500,00/R$ 3.000,00 - Nacionais/Importados."
            '''                    li = li + 1
            '''                Case 124
            '''                    pTexto(li) = "Cobertura 124 - R C Concessionárias - até 200 Km:  Nos sinistros ocorridos com veículos embarcados em guinchos/cegonhas de terceiros,"
            '''                    li = li + 1
            '''                    pTexto(li) = "a Franquia fica substituída pela participação obrigatória do segurado em 10% prej. mín. R$ 2.000,00/R$ 3.000,00 - Nacionais/Importados."
            '''                    li = li + 1
            '''                Case 125
            '''                    pTexto(li) = "Cobertura 125 - R C Concessionárias - até 300 Km:  Nos sinistros ocorridos com veículos embarcados em guinchos/cegonhas de terceiros,"
            '''                    li = li + 1
            '''                    pTexto(li) = "a Franquia fica substituída pela participação obrigatória do segurado em 10% prej. mín. R$ 2.000,00/R$ 3.000,00 - Nacionais/Importados."
            '''                    li = li + 1
            '''            End Select
        Next lobjstcCobert
    End If

    If gM46V111.gCod_plano = 2139 Then
        pTexto(lI) = " Este seguro destina-se a cobrir exclusivamente o conteúdo do imóvel."
        lI = lI + 1
    End If

End Sub
Private Function mfCriDados_Debito_Fonte12(ByVal pNumArqDados As Integer, ByVal pobjstcPed As stcA46V702B) As Boolean
'Função: cria arquivo de dados - autorização para débito em conta.

'Parâmetros de entrada...pNumArqDados.....Número do arquivo de dados.

    Dim lAuxiliar As String    'Auxiliar.
    Dim lLocData As String    'Local/Data.

    mfCriDados_Debito_Fonte12 = False

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
    If pobjstcPed.Num_PI_Emissao <> 0 Then
        lAuxiliar = lAuxiliar & Space(10) & "PI: " & pobjstcPed.Num_PI_Emissao & "-" & pobjstcPed.Dig_PI_Emissao
    Else
        lAuxiliar = lAuxiliar & Space(10) & "PROP. NÃO TRANSM."
    End If
    If pobjstcPed.Ind_Reentrada = 1 Then
        lAuxiliar = lAuxiliar & "**REENTRADA**"
    End If
    Print #pNumArqDados, lAuxiliar
    'CNPJ - CNPJ/CPF.
    Print #pNumArqDados, pobjstcPed.gCNPJ_CPF_Formatado
    'PROP - Número do pedido/proposta.
    Print #pNumArqDados, pobjstcPed.NOSSO_NUMERO & " DV: " & pobjstcPed.Cod_Unicidade & "-" & pobjstcPed.Cod_Ver_Unici

    '3. Banco...
''20170123: Ajustado para buscar descricao dos bancos de debito do banco de dados (TBM_BANCOS)
''pobjstcPed.gDesc_Banco passa a ser uma property q busca a descricao no banco.
'    Select Case pobjstcPed.Cod_Bco_Cobr
'    Case 1    'Banco do Brasil  'Projeto 11080 - Debito automatico - Banco do Brasil
'        Print #pNumArqDados, "Banco do Brasil"
'    Case 347, 356  'Sudameris.
'        Print #pNumArqDados, "Real"
'    Case 479    'Bank Boston.
'        Print #pNumArqDados, "Bank Boston"
'    Case 237    'Bradesco.
'        Print #pNumArqDados, "Bradesco"
'    Case 341    'Banco Itaú.
'        Print #pNumArqDados, "Itaú"
'    Case 409    'Unibanco.
'        Print #pNumArqDados, "Unibanco"
'    Case 399    'HSBC
'        Print #pNumArqDados, "HSBC"
'    Case 33     'Santander
'        Print #pNumArqDados, "Santander"
'    Case 47     'Banese
'        Print #pNumArqDados, "Banese"
'    Case Else
'        Print #pNumArqDados, " -"
'    End Select
    
    Print #pNumArqDados, pobjstcPed.gDesc_Banco

    'AGEN - Agência.
    lAuxiliar = ""
    lAuxiliar = pobjstcPed.gAgenciaDig
    If pobjstcPed.Cod_Bco_Cobr = 237 Or pobjstcPed.Cod_Bco_Cobr = 1 Then    'Projeto 11080 - Debito automatico - Banco do Brasil
        Print #pNumArqDados, Left(Format(lAuxiliar, "00000"), 4) & "-" & Right(lAuxiliar, 1)
    Else
        Print #pNumArqDados, lAuxiliar
    End If
    'CCOR - Conta corrente.
    lAuxiliar = ""
    Call mpForConta1(pobjstcPed.Cod_Bco_Cobr, pobjstcPed.mNum_Cta_Bas_Cor_Compl, lAuxiliar)
    Print #pNumArqDados, lAuxiliar
    'LINH - Linha sobre as parcelas.
    If val(pobjstcPed.objstcCtrlEmis.Num_Bloq & pobjstcPed.objstcCtrlEmis.DIG_BLOQUETO) <> 0 Then
        Print #pNumArqDados, Format(pobjstcPed.Qtd_Parc, "00") - 1 & " parcelas de R$ " & _
                             Format(pobjstcPed.val_DemParcelas, "Standard") & " cada."
    Else
        Print #pNumArqDados, "1ª Parcela de R$ " & Format(pobjstcPed.Val_PriParcela, "Standard") & " + " & _
                             Format(pobjstcPed.Qtd_Parc, "00") - 1 & " parcelas de R$ " & _
                             Format(pobjstcPed.val_DemParcelas, "Standard") & " cada."
    End If
    'DATA - Local e data.
    Print #pNumArqDados, lLocData


    mfCriDados_Debito_Fonte12 = True
End Function


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
    Case 347, 356  'Sudameris.  'Real
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
    If val(pobjstcPed.objstcCtrlEmis.Num_Bloq & pobjstcPed.objstcCtrlEmis.DIG_BLOQUETO) <> 0 Then
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
    Dim clsDigito As clsA99V001
    Dim lDigito As String
    Dim lMensagem As String
    Dim lConta As String

    Select Case pCodBanco
    Case 1    'Banco do Brasil. 'Projeto 11080 - Debito automatico - Banco do Brasil
        lAuxiliar = Format$(pNumConta, "00000000")
        If Mid$(lAuxiliar, 8, 1) = "0" Then
            Set clsDigito = New clsA99V001
            Set clsDigito.M46V999 = gM46V999

            Call clsDigito.DIG_BANCOBRASIL(lAuxiliar, lDigito, lMensagem)
            If lDigito = "X" Then lAuxiliar = Mid$(lAuxiliar, 1, 7) & "X"
        End If
        pSaida = Mid$(lAuxiliar, 1, 7) & "-" & Mid$(lAuxiliar, 8, 1)
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
    Case 356    'Real
        lAuxiliar = Format$(pNumConta, "0000000000")
        pSaida = Mid$(lAuxiliar, 1, 9) & "-" & Mid$(lAuxiliar, 10, 1)
    Case 399
        lAuxiliar = Format$(pNumConta, "00000000000")
        pSaida = Mid$(lAuxiliar, 1, 9) & "-" & Mid$(lAuxiliar, 10, 2)
    Case 33
        lAuxiliar = Format$(pNumConta, "000000000")
        pSaida = Mid$(lAuxiliar, 1, 8) & "-" & Mid$(lAuxiliar, 9, 1)
    Case 47
        lAuxiliar = Format$(pNumConta, "000000000")
        pSaida = Mid$(lAuxiliar, 1, 8) & "-" & Mid$(lAuxiliar, 9, 1)
    Case Else       'Outros bancos.
        pSaida = Format$(pNumConta, "0000000000")
    End Select
End Sub

Private Sub mpDescricaoFranquia(pobjstcPedLoc As stcA46V708B, pobjstcPed As stcA46V702B, _
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
    Dim lRegiao As Integer
    Dim lSelect As String

    MsgBox "ROTINA DESATIVADA !!!"

    If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, _
                              lBDP00Multi, lMensagem) = False Then
        Call gM46V999.gpGraLog(1, lMensagem)
        Exit Sub
    End If

    If pobjstcPed.ColecaoPedCorr.Count > 0 Then
        Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(1)
        lCorretor = lObjstcCorr.Cod_Corr
    Else
        lCorretor = 0
    End If

    'Ficha 901383
    'IF gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then

    'Consulta por todos os corretores + todas as classificações
    lSelect = " Select * from TAB_DESC_FRANQUIA Where"
    lSelect = lSelect & " Cod_Ramo = " & pobjstcPed.Cod_Ramo
    lSelect = lSelect & " and Cod_Produto = " & pobjstcPedLoc.Cod_Produto
    lSelect = lSelect & "  and (Cod_Regiao = 0 or Cod_Regiao = " & pobjstcPedLoc.Cod_Regiao & ")"
    lSelect = lSelect & "  and (Compl_Clasf ='0')"
    lSelect = lSelect & "  and (Cod_Corretor = 0)"
    lSelect = lSelect & "  and (Cod_Cobert not in (Select Cod_Cobert from TAB_DESC_FRANQUIA where Cod_Corretor = " & lCorretor & "))"
    lSelect = lSelect & " and Cod_Tarf_Ini <= " & pobjstcPed.Cod_Produto
    lSelect = lSelect & " and Cod_Tarf_Final >= " & pobjstcPed.Cod_Produto
    lSelect = lSelect & "  order by Cod_Corretor desc,Cod_Cobert,cod_produto,Compl_Clasf desc, Qtde_Linha"

    If gM46V999.gfObtRegistro(lBDP00Multi, lSelect, lrst, lMensagem) = False Then
        Call gM46V999.gpGraLog(1, lMensagem)
        ReDim pTextos(6)
        Exit Sub
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

    'Ficha 1000284
    Dim lValFranquia As Double

    Do While Not lrst.EOF

        For Each lobjstcCobert In pobjstcPedLoc.ColecaoPedCobert
            lValFranquia = 0
            ' Call mpBuscaLimFranquia_Cobert(lobjstcCobert.Cod_Cobert, lValFranquia)
            MsgBox "ROTINA DESATIVADA - AVISAR HELPDESK "

            If lValFranquia > 0 And lValFranquia <> lobjstcCobert.Val_Franq_Min Then
                GoTo Pula_Texto
            End If

            If lobjstcCobert.Cod_Cobert = lrst("Cod_Cobert") Then
                Select Case lobjstcCobert.Cod_Cobert
                Case gM46V999.gfBuscarDePara(39, 1), gM46V999.gfBuscarDePara(40, 1), gM46V999.gfBuscarDePara(41, 1), gM46V999.gfBuscarDePara(42, 1), gM46V999.gfBuscarDePara(43, 1), _
                gM46V999.gfBuscarDePara(44, 1), gM46V999.gfBuscarDePara(47, 1), gM46V999.gfBuscarDePara(50, 1)
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
                    lTextos(lI) = lrst("Descricao")
                    lCount = lCount + 1
                End If
                lCod_Ant = lrst("Cod_Cobert")
                lCod_Linha = lrst("Qtde_Linha")
                lI = lI + 1
                If lI = pTamArray Then Exit Do
                Exit For
            End If
Pula_Texto:
        Next lobjstcCobert
        lrst.MoveNext
    Loop

    If lTextoFranquia = True Then
        lTextos(lI) = "Franquia: 7 (sete) dias"
        lI = lI + 1
    End If
    Call gM46V999.gpFecha2(lBDP00Multi)

FinalTextos:
    pTextos = lTextos
End Sub
Private Sub mpDescricaoFranquia_CorretorCoberturaDesativada(pobjstcPedLoc As stcA46V708B, pobjstcPed As stcA46V702B, pCodCobert As Integer, _
                                                            pValoFranquia As Double, pTextos() As String, pTamArray As Integer)

    Dim lBDP00Multi As ADODB.Connection
    Dim lColecaoTemp As Collection
    Dim lI As Integer
    Dim lTextoFranquia As Boolean
    Dim lMensagem As String
    Dim lObjstcCorr As stcA46V705B    'Objeto stcA46V705B  (corretor).
    Dim lrst As clsYasRecordSet
    Dim lTextos() As String
    Dim lCod_Ant As Integer
    Dim lCod_Linha As Integer
    Dim lCount As Integer
    Dim lCorretor As Double
    Dim lCobertura As Integer
    Dim lRegiao As Integer
    Dim lSelect As String

    On Error GoTo ControleErro

    If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, _
                              lBDP00Multi, lMensagem) = False Then
        Call gM46V999.gpGraLog(1, lMensagem)
        Exit Sub
    End If

    If pobjstcPed.ColecaoPedCorr.Count > 0 Then
        Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(1)
        lCorretor = lObjstcCorr.Cod_Corr
    Else
        lCorretor = 0
    End If

    'Ficha 1000655
    lRegiao = 0   'vai estar errado

    'Vai procurar na tabela de Ajustes pontuais se tem uma região definida com franquia diferenciada
    If gM46V999.gfFranquiaRegiao(e_Tip_Valor_Regiao, TpoValor, pobjstcPed.Cod_Ramo, pobjstcPed.Cod_Estip, pobjstcPedLoc.Cod_Clasf, pobjstcPedLoc.Cod_Compl_Clasf, pCodCobert, pobjstcPed.Cod_Produto, pobjstcPedLoc.Cod_Constr, pobjstcPedLoc.Num_Cep_Risco, lRegiao, lMensagem) = False Then
        Exit Sub
    End If

    'Ficha 901383
    'IF gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then

    'Consulta por Corretor + Classificação
    lSelect = " Select * from TAB_DESC_FRANQUIA Where"
    lSelect = lSelect & " Cod_Ramo = " & pobjstcPed.Cod_Ramo
    lSelect = lSelect & " and Cod_Produto = " & pobjstcPedLoc.Cod_Produto
    lSelect = lSelect & "  and (Cod_Regiao = 0 or Cod_Regiao = " & lRegiao & ")"
    lSelect = lSelect & "  and Compl_Clasf = '" & Trim(pobjstcPedLoc.Cod_Clasf) & Trim(pobjstcPedLoc.Cod_Compl_Clasf) & "'"
    lSelect = lSelect & "  and (Cod_Corretor = " & lCorretor & " or Cod_Corretor = 0)"
    lSelect = lSelect & " and Cod_Tarf_Ini <= " & pobjstcPed.Cod_Produto
    lSelect = lSelect & " and Cod_Tarf_Final >= " & pobjstcPed.Cod_Produto
    lSelect = lSelect & " and Cod_Cobert = " & pCodCobert
    lSelect = lSelect & "  order by Cod_Corretor desc,Cod_Cobert,cod_produto,Compl_Clasf desc, Qtde_Linha"

    If gM46V999.gfObtRegistro(lBDP00Multi, lSelect, lrst, lMensagem) = False Then
        Call gM46V999.gpGraLog(1, lMensagem)
        ReDim pTextos(6)
        Exit Sub
    End If

    If lrst.EOF Then
        'Consulta por Corretor + todas as classificações
        lSelect = " Select * from TAB_DESC_FRANQUIA Where"
        lSelect = lSelect & " Cod_Ramo = " & pobjstcPed.Cod_Ramo
        lSelect = lSelect & " and Cod_Produto = " & pobjstcPedLoc.Cod_Produto
        lSelect = lSelect & "  and (Cod_Regiao = 0 or Cod_Regiao = " & lRegiao & ")"
        lSelect = lSelect & "  and (Compl_Clasf ='0')"
        lSelect = lSelect & "  and (Cod_Corretor = " & lCorretor & ")"
        lSelect = lSelect & " and Cod_Tarf_Ini <= " & pobjstcPed.Cod_Produto
        lSelect = lSelect & " and Cod_Tarf_Final >= " & pobjstcPed.Cod_Produto
        lSelect = lSelect & " and Cod_Cobert = " & pCodCobert
        lSelect = lSelect & "  order by Cod_Corretor desc,Cod_Cobert,cod_produto,Compl_Clasf desc, Qtde_Linha"

        If gM46V999.gfObtRegistro(lBDP00Multi, lSelect, lrst, lMensagem) = False Then
            Call gM46V999.gpGraLog(1, lMensagem)
            ReDim pTextos(6)
            Exit Sub
        End If
    End If

    If lrst.EOF Then
        'Consulta por cobertura
        lSelect = " Select * from TAB_DESC_FRANQUIA Where"
        lSelect = lSelect & " Cod_Ramo = " & pobjstcPed.Cod_Ramo
        lSelect = lSelect & " and Cod_Produto = " & pobjstcPedLoc.Cod_Produto
        lSelect = lSelect & "  and (Cod_Regiao = 0 or Cod_Regiao = " & lRegiao & ")"
        lSelect = lSelect & "  and (Compl_Clasf ='0')"
        lSelect = lSelect & "  and (Cod_Corretor = 0 )"
        lSelect = lSelect & " and Cod_Tarf_Ini <= " & pobjstcPed.Cod_Produto
        lSelect = lSelect & " and Cod_Tarf_Final >= " & pobjstcPed.Cod_Produto
        lSelect = lSelect & " and Cod_Cobert = " & pCodCobert
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
    End If

    Dim strTexto As String
    strTexto = ""

    'Ficha 1000284
    Dim lValFranquia As Double

    Do While Not lrst.EOF
        lValFranquia = 0

        'Call mpBuscaLimFranquia_Cobert(lrst("Cod_Cobert"), lValFranquia)

        'Call mpProcuraValorDefault(pobjstcPedLoc, pobjstcPed, lrst("Cod_Cobert"), lValFranquia)

        If lValFranquia > 0 And lValFranquia <> pValoFranquia Then
            GoTo Pula_Texto
        End If

        Select Case lrst("Cod_Cobert")
        Case gM46V999.gfBuscarDePara(39, 1), gM46V999.gfBuscarDePara(40, 1), gM46V999.gfBuscarDePara(41, 1), gM46V999.gfBuscarDePara(42, 1), gM46V999.gfBuscarDePara(43, 1), _
        gM46V999.gfBuscarDePara(44, 1), gM46V999.gfBuscarDePara(47, 1), gM46V999.gfBuscarDePara(50, 1)
            If pobjstcPed.Cod_Estip < 1020 Then
                lTextoFranquia = True
            End If
        End Select
        If lCod_Linha + 1 <> lrst("Qtde_Linha") Then
            Exit Do
        End If
        If lCod_Ant = lrst("Cod_Cobert") Then
            '''                lTextos(li) = "   " & lrst("Descricao")
        Else
            '''                lTextos(li) = lrst("Descricao")
            lCount = lCount + 1
        End If
        strTexto = strTexto & IIf(strTexto <> "", " ", "") & lrst("Descricao")
        lCod_Ant = lrst("Cod_Cobert")
        lCod_Linha = lrst("Qtde_Linha")
        '''            li = li + 1
        '''            If li = pTamArray Then Exit Do

Pula_Texto:
        lrst.MoveNext
    Loop

    If strTexto <> "" Then
        lTextos(lI) = strTexto
        lI = lI + 1
    End If

    If lTextoFranquia = True Then
        lTextos(lI) = "Franquia: 7 (sete) dias"
        lI = lI + 1
    End If

    'Rotina para ajustar quebra do texto no array de franquia
    Call mpQuebraTextoFranquia(lTextos, 110)

    Call gM46V999.gpFecha2(lBDP00Multi)

FinalTextos:
    pTextos = lTextos
    Exit Sub
ControleErro:
    Call gM46V999.gpGraLog(1, Err.Description)
End Sub


Private Sub mpDescricaoFranquia_CorretorDesativado(pobjstcPedLoc As stcA46V708B, pobjstcPed As stcA46V702B, _
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
    Dim lRegiao As Integer
    Dim lSelect As String

    On Error GoTo ControleErro

    If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, _
                              lBDP00Multi, lMensagem) = False Then
        Call gM46V999.gpGraLog(1, lMensagem)
        Exit Sub
    End If

    If pobjstcPed.ColecaoPedCorr.Count > 0 Then
        Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(1)
        lCorretor = lObjstcCorr.Cod_Corr
    Else
        lCorretor = 0
    End If

    'Ficha 901383
    'IF gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then
    lI = 0
    lTextoFranquia = False
    ReDim lTextos(pTamArray)
    lCount = 1
    For Each lobjstcCobert In pobjstcPedLoc.ColecaoPedCobert

        'Ficha 1000655
        lRegiao = 0    'tem que estar ok pq pega a cobertura do objeto

        'Vai procurar na tabela de Ajustes pontuais se tem uma região definida com franquia diferenciada
        If gM46V999.gfFranquiaRegiao(e_Tip_Valor_Regiao, TpoValor, pobjstcPed.Cod_Ramo, pobjstcPed.Cod_Estip, pobjstcPedLoc.Cod_Clasf, pobjstcPedLoc.Cod_Compl_Clasf, lobjstcCobert.Cod_Cobert, pobjstcPed.Cod_Produto, pobjstcPedLoc.Cod_Constr, pobjstcPedLoc.Num_Cep_Risco, lRegiao, lMensagem) = False Then
            Exit Sub
        End If

        'Consulta por Corretor + Classificação
        lSelect = " Select * from TAB_DESC_FRANQUIA Where"
        lSelect = lSelect & " Cod_Ramo = " & pobjstcPed.Cod_Ramo
        lSelect = lSelect & " and Cod_Produto = " & pobjstcPedLoc.Cod_Produto
        lSelect = lSelect & "  and (Cod_Regiao = 0 or Cod_Regiao = " & lRegiao & ")"
        lSelect = lSelect & "  and Compl_Clasf = '" & Trim(pobjstcPedLoc.Cod_Clasf) & Trim(pobjstcPedLoc.Cod_Compl_Clasf) & "'"
        lSelect = lSelect & "  and (Cod_Corretor = " & lCorretor & " or Cod_Corretor = 0)"
        lSelect = lSelect & " and Cod_Tarf_Ini <= " & pobjstcPed.Cod_Produto
        lSelect = lSelect & " and Cod_Tarf_Final >= " & pobjstcPed.Cod_Produto
        lSelect = lSelect & " and cod_cobert = " & lobjstcCobert.Cod_Cobert
        lSelect = lSelect & "  order by Cod_Corretor desc,Cod_Cobert,cod_produto,Compl_Clasf desc, Qtde_Linha"

        If gM46V999.gfObtRegistro(lBDP00Multi, lSelect, lrst, lMensagem) = False Then
            Call gM46V999.gpGraLog(1, lMensagem)
            ReDim pTextos(6)
            Exit Sub
        End If

        If lrst.EOF Then
            'Consulta por Corretor + todas as classificações
            lSelect = " Select * from TAB_DESC_FRANQUIA Where"
            lSelect = lSelect & " Cod_Ramo = " & pobjstcPed.Cod_Ramo
            lSelect = lSelect & " and Cod_Produto = " & pobjstcPedLoc.Cod_Produto
            lSelect = lSelect & "  and (Cod_Regiao = 0 or Cod_Regiao = " & lRegiao & ")"
            lSelect = lSelect & "  and (Compl_Clasf ='0')"
            lSelect = lSelect & "  and (Cod_Corretor = " & lCorretor & ")"
            lSelect = lSelect & " and Cod_Tarf_Ini <= " & pobjstcPed.Cod_Produto
            lSelect = lSelect & " and Cod_Tarf_Final >= " & pobjstcPed.Cod_Produto
            lSelect = lSelect & " and cod_cobert = " & lobjstcCobert.Cod_Cobert
            lSelect = lSelect & "  order by Cod_Corretor desc,Cod_Cobert,cod_produto,Compl_Clasf desc, Qtde_Linha"

            If gM46V999.gfObtRegistro(lBDP00Multi, lSelect, lrst, lMensagem) = False Then
                Call gM46V999.gpGraLog(1, lMensagem)
                ReDim pTextos(6)
                Exit Sub
            End If
        End If


        If lrst.EOF Then
            'Consulta por todos os corretores + todas as classificações
            lSelect = " Select * from TAB_DESC_FRANQUIA Where"
            lSelect = lSelect & " Cod_Ramo = " & pobjstcPed.Cod_Ramo
            lSelect = lSelect & " and Cod_Produto = " & pobjstcPedLoc.Cod_Produto
            lSelect = lSelect & "  and (Cod_Regiao = 0 or Cod_Regiao = " & lRegiao & ")"
            lSelect = lSelect & "  and (Compl_Clasf ='0')"
            lSelect = lSelect & "  and (Cod_Corretor = 0)"
            lSelect = lSelect & " and Cod_Tarf_Ini <= " & pobjstcPed.Cod_Produto
            lSelect = lSelect & " and Cod_Tarf_Final >= " & pobjstcPed.Cod_Produto
            lSelect = lSelect & " and cod_cobert = " & lobjstcCobert.Cod_Cobert
            lSelect = lSelect & "  order by Cod_Corretor desc,Cod_Cobert,cod_produto,Compl_Clasf desc, Qtde_Linha"

            If gM46V999.gfObtRegistro(lBDP00Multi, lSelect, lrst, lMensagem) = False Then
                Call gM46V999.gpGraLog(1, lMensagem)
                ReDim pTextos(6)
                Exit Sub
            End If
        End If

        Dim strTexto As String
        strTexto = ""

        'Ficha 1000284
        Dim lValFranquia As Double

        Do While Not lrst.EOF
            lValFranquia = 0

            'Call mpBuscaLimFranquia_Cobert(lobjstcCobert.Cod_Cobert, lValFranquia)
            'Call mpProcuraValorDefault(pobjstcPedLoc, pobjstcPed, lobjstcCobert.Cod_Cobert, lValFranquia)

            '1001135 - Corretor Volks - Franquia Especial Concessionaria
            If gM46V999.gCodProduto >= 1002 Then
                If (InStr(1, "00008001/00008002/00008003/00008004/00008005/00008006/00008008/00008011/00008014/00008028", Format(frmT46V102A.txtCorretor(1).Tag, "00000000"), vbTextCompare) > 0) And _
                   (val(pobjstcPedLoc.Cod_Clasf) = gM46V999.gfBuscarDePara(22, 2) And val(pobjstcPedLoc.Cod_Compl_Clasf) = gM46V999.gfBuscarDePara(22, 2, 13)) And _
                   (InStr(1, "0038/0064/0066", Format(lobjstcCobert.Cod_Cobert, "0000"), vbTextCompare) > 0) Then
                    GoTo Pula_Texto
                Else
                    If lValFranquia > 0 And lValFranquia <> lobjstcCobert.Val_Franq_Min Then
                        GoTo Pula_Texto
                    End If
                End If
            Else
                If lValFranquia > 0 And lValFranquia <> lobjstcCobert.Val_Franq_Min Then
                    GoTo Pula_Texto
                End If
            End If

            Select Case lobjstcCobert.Cod_Cobert
            Case gM46V999.gfBuscarDePara(39, 1), gM46V999.gfBuscarDePara(40, 1), gM46V999.gfBuscarDePara(41, 1), gM46V999.gfBuscarDePara(42, 1), gM46V999.gfBuscarDePara(43, 1), _
            gM46V999.gfBuscarDePara(44, 1), gM46V999.gfBuscarDePara(47, 1), gM46V999.gfBuscarDePara(50, 1)
                If pobjstcPed.Cod_Estip < 1020 Then
                    lTextoFranquia = True
                End If
            End Select
            If lCod_Ant = lobjstcCobert.Cod_Cobert And _
               lCod_Linha + 1 <> lrst("Qtde_Linha") Then
                Exit For
            End If
            If lCod_Ant = lobjstcCobert.Cod_Cobert Then
                ''''                lTextos(li) = "   " & lrst("Descricao")
            Else
                ''''                lTextos(li) = lrst("Descricao")
                lCount = lCount + 1
            End If
            strTexto = strTexto & IIf(strTexto <> "", " ", "") & lrst("Descricao")
            lCod_Ant = lrst("Cod_Cobert")
            lCod_Linha = lrst("Qtde_Linha")
            '''            li = li + 1
            '''            If li = pTamArray Then Exit Do

Pula_Texto:
            lrst.MoveNext
        Loop

        If strTexto <> "" Then
            lTextos(lI) = strTexto
            lI = lI + 1
        End If

        If lTextoFranquia = True Then
            lTextos(lI) = "Franquia: 7 (sete) dias"
            lI = lI + 1
        End If

    Next lobjstcCobert

    'Rotina para ajustar quebra do texto no array de franquia
    Call mpQuebraTextoFranquia(lTextos, 110)

    Call gM46V999.gpFecha2(lBDP00Multi)
FinalTextos:
    pTextos = lTextos
    Exit Sub
ControleErro:
    Call gM46V999.gpGraLog(1, Err.Description)
End Sub

Private Function gCoefComissao(pComissao As Single) As Single
    gCoefComissao = (1 - gM46O001.gProComBasica / 100) / (1 - pComissao / 100)
End Function

Private Sub mpPrazoTabCoefPI(pobjstcPed As stcA46V702B, pCodCobert As Integer, _
                             pNumMeses As Integer, pPrazo As Integer)
    Dim lBDP00Multi As ADODB.Connection
    Dim lMensagem As String
    Dim lrst As clsYasRecordSet
    Dim lSelect As String

    If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, _
                              lBDP00Multi, lMensagem) = False Then
        Call gM46V999.gpGraLog(1, lMensagem)
        Exit Sub
    End If
    lSelect = " Select Prazo from Tbm_COEF_PI Where"
    lSelect = lSelect & " Cod_Ramo = " & pobjstcPed.Cod_Ramo
    lSelect = lSelect & " and Cod_Cobert = " & pCodCobert
    lSelect = lSelect & " and Num_meses = " & pNumMeses
    lSelect = lSelect & " and Cod_Tarif = " & pobjstcPed.Cod_Produto
    lSelect = lSelect & " and Cod_Sequencia <= " & pobjstcPed.Cod_Sequencia

    If gM46V999.gfObtRegistro(lBDP00Multi, lSelect, lrst, lMensagem) = False Then
        Call gM46V999.gpGraLog(1, lMensagem)
        pPrazo = 0
        Exit Sub
    End If
    pPrazo = 0
    If Not lrst.EOF Then
        pPrazo = IIf(IsNull(lrst("Prazo")), 0, lrst("Prazo"))
    End If

    'Supermercado SMART
    If pobjstcPed.Cod_Ramo = 112 And pobjstcPed.Cod_Produto >= 1042 And (pobjstcPed.Cod_Estip >= 2178 And pobjstcPed.Cod_Estip <= 2182) And pCodCobert = 39 Then
        pPrazo = 7
    End If
    Call gM46V999.gpFecha2(lBDP00Multi)

End Sub

Private Function mfCriDados_MULTCOTA(ByVal pNumArqDados As Integer, ByVal pobjstcPed As stcA46V702B, _
                                     pobjstcPedLoc As stcA46V708B) As Boolean

'Função: cria arquivo de dados - cotação empresarial simples Anexo.

'Parâmetro de entrada...pNumArqDados...Número do arquivo de dados.


    Dim lAuxiliar1 As Double  'Auxiliar - utilizado apenas como exemplo.
    Dim lAuxiliar2 As String  'Auxiliar - utilizado apenas como exemplo.
    Dim lAuxiliar3 As String  'Auxiliar - utilizado apenas como exemplo.
    Dim lObjstcCorr As stcA46V705B    'Objeto stcA46V705B  (corretor).
    Dim lCountLine As Integer    'Contador de linhas.
    Dim lobjstcBem As stcA46V715B
    Dim lMensagem As String
    Dim lbdTab_Tipo_Fator As ADODB.Connection   'Arquivo: P00Tarifa.
    Dim lrsTab_Tipo_Fator As clsYasRecordSet    'Registro :Tab_Tipo_Fator.
    Dim lobjstcFator As stcA46V720B   'Objeto da classe de estrutura stcA46V720B (TAB_PED_FAT_RISCO).
    Dim lblnAchou As Boolean    'Indica se achou  o fator ou não.
    Dim lQuant As Integer
    Dim lTotalPag As Integer
    Dim lTemInfo As Boolean
    Dim lSelect As String

    mfCriDados_MULTCOTA = False
    lQuant = 0
    lTemInfo = False
    If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTab_Tipo_Fator, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Function
    End If

    lQuant = 1
    lTotalPag = 1
    '001. Título do seguro (de acordo com a base de dados).
    '     Neste caso deve ser colocado YASUDA RESIDENCIAL, YASUDA CONDOMÍNIO ou YASUDA EMPRESARIAL.
    lAuxiliar2 = "YASUDA " & UCase(pobjstcPed.gDesc_Ramo)

    Do While Len(lAuxiliar2) < 45
        lAuxiliar2 = " " & lAuxiliar2 & " "
    Loop

    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)



    '002. Página.
    Print #pNumArqDados, "Página: " & lQuant & "/" & lTotalPag

    '003. Emissão.
    Print #pNumArqDados, "Emissão: " & Format$(Now, "dd/mm/yyyy hh:nn:ss")

    'Ficha 901223
    '001B. Código do Processo Susep.
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


    '004. Validade.
    If pobjstcPedLoc.Tip_Emissao = e_TipEmissao_RenYasuda Then
        If gM46V999.gCodProduto >= 978 And gM46V999.gCodProduto <= 1047 Then
            Print #pNumArqDados, "Validade: " & Format$(CDate(gM46V999.gfForData(2, pobjstcPed.Dat_Pri_Calculo)) + 15, "dd/mm/yyyy")
        Else
            Print #pNumArqDados, "Validade: " & Format$(CDate(gM46V999.gfForData(2, pobjstcPed.Dat_Pri_Calculo)) + 30, "dd/mm/yyyy")
        End If
    Else
        Print #pNumArqDados, "Validade: " & Format$(CDate(gM46V999.gfForData(2, pobjstcPed.Dat_Pri_Calculo)) + 15, "dd/mm/yyyy")
    End If

    '005. Número do pedido .
    lAuxiliar1 = pobjstcPed.Num_Ped
    Print #pNumArqDados, Format$(lAuxiliar1, "000000")

    '006. Número do item .
    lAuxiliar1 = pobjstcPedLoc.Num_Item
    Print #pNumArqDados, Format$(lAuxiliar1, "0000")

    '007. Corretor (código e nome) .
    If pobjstcPed.ColecaoPedCorr.Count > 0 Then
        Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(1)
        lAuxiliar1 = lObjstcCorr.Cod_Corr
        lAuxiliar2 = lObjstcCorr.gNom_Corretor
    Else
        lAuxiliar1 = 0
        lAuxiliar2 = " "
    End If
    Print #pNumArqDados, Format$(lAuxiliar1, "0000") & " - " & lAuxiliar2

    '008. Tarifa .
    lAuxiliar2 = pobjstcPed.Cod_Produto
    If pobjstcPed.Cod_Produto = 976 And pobjstcPed.Cod_Ramo = e_CodRamo_Empresarial Then
        lAuxiliar2 = lAuxiliar2 & " A"
    End If
    If pobjstcPed.Cod_Sequencia = 1 Then
        lAuxiliar2 = lAuxiliar2 & " (A)"
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '009. Nosso número .
    lAuxiliar2 = pobjstcPed.NOSSO_NUMERO
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '010 a 024. Fatores de Proteção - linhas de 01 a 25.

    'Ficha 1000601
    If gM46V999.gCod_Ramo = e_CodRamo_Residencial Then
        If pobjstcPedLoc.gLiberaDescontoEspecial = True Then
            lSelect = " Select distinct DSC_FATOR, COD_FATOR  from Tbm_Tipo_Fator_Protec "
            lSelect = lSelect & " Where cod_Ramo = " & pobjstcPed.Cod_Ramo
            lSelect = lSelect & " and Cod_Modalidade = " & pobjstcPedLoc.Cod_Produto
            lSelect = lSelect & " and COD_TARIF = " & pobjstcPed.Cod_Produto
            lSelect = lSelect & " and cod_Sequencia  <= " & pobjstcPed.Cod_Sequencia
            lSelect = lSelect & " order by DSC_FATOR"
        Else
            lSelect = " Select distinct DSC_FATOR, COD_FATOR  from Tbm_Tipo_Fator_Protec "
            lSelect = lSelect & " Where cod_Ramo = " & pobjstcPed.Cod_Ramo
            lSelect = lSelect & " and Cod_Modalidade = " & pobjstcPedLoc.Cod_Produto
            lSelect = lSelect & " and COD_TARIF = " & pobjstcPed.Cod_Produto
            lSelect = lSelect & " and cod_Sequencia  <= " & pobjstcPed.Cod_Sequencia
            lSelect = lSelect & " and Cod_Fator <> 7 "    'Não carregar o item Desconto Especial
            lSelect = lSelect & " order by DSC_FATOR"
        End If
    Else
        lSelect = " Select distinct DSC_FATOR, COD_FATOR  from Tbm_Tipo_Fator_Protec "
        lSelect = lSelect & " Where cod_Ramo = " & pobjstcPed.Cod_Ramo
        lSelect = lSelect & " and Cod_Modalidade = " & pobjstcPedLoc.Cod_Produto
        lSelect = lSelect & " and COD_TARIF = " & pobjstcPed.Cod_Produto
        lSelect = lSelect & " and cod_Sequencia  <= " & pobjstcPed.Cod_Sequencia
        lSelect = lSelect & " order by DSC_FATOR"
    End If

    If gM46V999.gfObtRegistro(lbdTab_Tipo_Fator, lSelect, lrsTab_Tipo_Fator, lMensagem) = False Then
        Call gM46V999.gpFecha2(lbdTab_Tipo_Fator)
        Call gM46V999.gpGraLog(2, lMensagem)
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    lCountLine = 0
    Do While Not lrsTab_Tipo_Fator.EOF
        lCountLine = lCountLine + 1
        lblnAchou = False
        For Each lobjstcFator In pobjstcPedLoc.ColecaoPedFator
            If lobjstcFator.Cod_Fator = lrsTab_Tipo_Fator("Cod_Fator") Then
                lTemInfo = True
                lblnAchou = True
                lAuxiliar2 = Format(lCountLine, "00") & " (X) " & Trim(lrsTab_Tipo_Fator("Dsc_Fator"))
                If Trim(lobjstcFator.Dsc_Compl_Fator) <> "" Then
                    lAuxiliar2 = lAuxiliar2 & " - " & Trim(lobjstcFator.Dsc_Compl_Fator)
                End If
                Print #pNumArqDados, lAuxiliar2
                Exit For
            End If
        Next lobjstcFator
        If lblnAchou = False Then
            lAuxiliar2 = Format(lCountLine, "00") & " ( ) " & Trim(lrsTab_Tipo_Fator("Dsc_Fator"))
            Print #pNumArqDados, lAuxiliar2
        End If
        lrsTab_Tipo_Fator.MoveNext
    Loop
    Do While Not lCountLine > 24
        Print #pNumArqDados, " "
        lCountLine = lCountLine + 1
    Loop

    lCountLine = 0
    For Each lobjstcBem In pobjstcPedLoc.ColecaoPedLocBem
        lCountLine = lCountLine + 1
        lAuxiliar3 = lobjstcBem.gTipBem & " - " & lobjstcBem.Dsc_Bem
        lAuxiliar2 = Format(lCountLine, "000") & " - " & Left(lAuxiliar3, 100) & String(90 - Len(Left(lAuxiliar3, 90)), ".") & "R$ " & Format(lobjstcBem.Val_Bem, "standard")
        Print #pNumArqDados, lAuxiliar2
        lTemInfo = True
    Next lobjstcBem

    Do While lCountLine < 34
        lCountLine = lCountLine + 1
        Print #pNumArqDados, " "
    Loop


    Call gM46V999.gpFecha2(lbdTab_Tipo_Fator)

    If lTemInfo = True Then
        mfCriDados_MULTCOTA = True
    Else
        Call gM46V999.gpFecha2(lbdTab_Tipo_Fator)
    End If

    '    mfCriDados_MULTCOTA = True

End Function

Private Function mfCriDados_MULTCOTC(ByVal pNumArqDados As Integer, ByVal pobjstcPed As stcA46V702B, _
                                     pNumDe As Integer, _
                                     pNumAte As Integer) As Boolean


'Função: cria arquivo de dados - cotação empresarial simples Anexo.

'Parâmetro de entrada...pNumArqDados...Número do arquivo de dados.


    Dim lAuxiliar1 As Double  'Auxiliar - utilizado apenas como exemplo.
    Dim lAuxiliar2 As String  'Auxiliar - utilizado apenas como exemplo.
    Dim lAuxiliar3 As String  'Auxiliar - utilizado apenas como exemplo.
    Dim lObjstcCorr As stcA46V705B    'Objeto stcA46V705B  (corretor).
    Dim lCountLine As Integer    'Contador de linhas.
    Dim lobjstcBem As stcA46V715B
    Dim lMensagem As String
    Dim lbdTab_Tipo_Fator As ADODB.Connection   'Arquivo: P00Tarifa.
    Dim lrsTab_Tipo_Fator As clsYasRecordSet    'Registro :Tab_Tipo_Fator.
    Dim lobjstcFator As stcA46V720B   'Objeto da classe de estrutura stcA46V720B (TAB_PED_FAT_RISCO).
    Dim lblnAchou As Boolean    'Indica se achou  o fator ou não.
    Dim lQuant As Integer
    Dim lTotalPag As Integer
    Dim lTemInfo As Boolean
    Dim lobjstcPedLoc As stcA46V708B
    Dim lSelect As String

    mfCriDados_MULTCOTC = False

    lQuant = 0
    lTemInfo = False
    If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTab_Tipo_Fator, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Function
    End If

    For Each lobjstcPedLoc In pobjstcPed.ColecaoPedLoc
        lQuant = 1
        lTotalPag = 1
        '001. Título do seguro (de acordo com a base de dados).
        '     Neste caso deve ser colocado YASUDA RESIDENCIAL, YASUDA CONDOMÍNIO ou YASUDA EMPRESARIAL.
        lAuxiliar2 = "YASUDA " & UCase(pobjstcPed.gDesc_Ramo)

        Do While Len(lAuxiliar2) < 45
            lAuxiliar2 = " " & lAuxiliar2 & " "
        Loop

        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)



        '002. Página.
        Print #pNumArqDados, "Página: " & lQuant & "/" & lTotalPag

        '003. Emissão.
        Print #pNumArqDados, "Emissão: " & Format$(Now, "dd/mm/yyyy hh:nn:ss")

        'Ficha 901223
        '001B. Código do Processo Susep.
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


        '004. Validade.
        If lobjstcPedLoc.Tip_Emissao = e_TipEmissao_RenYasuda Then
            If gM46V999.gCodProduto >= 978 Then
                Print #pNumArqDados, "Validade: " & Format$(CDate(gM46V999.gfForData(2, pobjstcPed.Dat_Pri_Calculo)) + 15, "dd/mm/yyyy")
            Else
                Print #pNumArqDados, "Validade: " & Format$(CDate(gM46V999.gfForData(2, pobjstcPed.Dat_Pri_Calculo)) + 30, "dd/mm/yyyy")
            End If
        Else
            Print #pNumArqDados, "Validade: " & Format$(CDate(gM46V999.gfForData(2, pobjstcPed.Dat_Pri_Calculo)) + 15, "dd/mm/yyyy")
        End If

        '005. Número do pedido .
        lAuxiliar1 = pobjstcPed.Num_Ped
        Print #pNumArqDados, Format$(lAuxiliar1, "000000")

        '006. Número do item .
        lAuxiliar1 = lobjstcPedLoc.Num_Item
        Print #pNumArqDados, Format$(lAuxiliar1, "0000")

        '007. Corretor (código e nome) .
        If pobjstcPed.ColecaoPedCorr.Count > 0 Then
            Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(1)
            lAuxiliar1 = lObjstcCorr.Cod_Corr
            lAuxiliar2 = lObjstcCorr.gNom_Corretor
        Else
            lAuxiliar1 = 0
            lAuxiliar2 = " "
        End If
        Print #pNumArqDados, Format$(lAuxiliar1, "0000") & " - " & lAuxiliar2

        '008. Tarifa .
        lAuxiliar2 = pobjstcPed.Cod_Produto
        If pobjstcPed.Cod_Produto = 976 And pobjstcPed.Cod_Ramo = e_CodRamo_Empresarial Then
            lAuxiliar2 = lAuxiliar2 & " A"
        End If
        If pobjstcPed.Cod_Sequencia = 1 Then
            lAuxiliar2 = lAuxiliar2 & " (A)"
        End If
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '009. Nosso número .
        lAuxiliar2 = pobjstcPed.NOSSO_NUMERO
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '010 a 024. Fatores de Proteção - linhas de 01 a 25.

        'Ficha 1000601
        If gM46V999.gCod_Ramo = e_CodRamo_Residencial Then
            If lobjstcPedLoc.gLiberaDescontoEspecial = True Then
                lSelect = " Select distinct DSC_FATOR, COD_FATOR  from Tbm_Tipo_Fator_Protec "
                lSelect = lSelect & " Where cod_Ramo = " & pobjstcPed.Cod_Ramo
                lSelect = lSelect & " and Cod_Modalidade = " & lobjstcPedLoc.Cod_Produto
                lSelect = lSelect & " and COD_TARIF = " & pobjstcPed.Cod_Produto
                lSelect = lSelect & " and cod_Sequencia  <= " & pobjstcPed.Cod_Sequencia
                lSelect = lSelect & " order by DSC_FATOR"
            Else
                lSelect = " Select distinct DSC_FATOR, COD_FATOR  from Tbm_Tipo_Fator_Protec "
                lSelect = lSelect & " Where cod_Ramo = " & pobjstcPed.Cod_Ramo
                lSelect = lSelect & " and Cod_Modalidade = " & lobjstcPedLoc.Cod_Produto
                lSelect = lSelect & " and COD_TARIF = " & pobjstcPed.Cod_Produto
                lSelect = lSelect & " and cod_Sequencia  <= " & pobjstcPed.Cod_Sequencia
                lSelect = lSelect & " and Cod_Fator <> 7 "    'Não carregar o item Desconto Especial
                lSelect = lSelect & " order by DSC_FATOR"
            End If
        Else
            lSelect = " Select distinct DSC_FATOR, COD_FATOR  from Tbm_Tipo_Fator_Protec "
            lSelect = lSelect & " Where cod_Ramo = " & pobjstcPed.Cod_Ramo
            lSelect = lSelect & " and Cod_Modalidade = " & lobjstcPedLoc.Cod_Produto
            lSelect = lSelect & " and COD_TARIF = " & pobjstcPed.Cod_Produto
            lSelect = lSelect & " and cod_Sequencia  <= " & pobjstcPed.Cod_Sequencia
            lSelect = lSelect & " order by DSC_FATOR"
        End If

        If gM46V999.gfObtRegistro(lbdTab_Tipo_Fator, lSelect, lrsTab_Tipo_Fator, lMensagem) = False Then
            Call gM46V999.gpFecha2(lbdTab_Tipo_Fator)
            Call gM46V999.gpGraLog(2, lMensagem)
            Screen.MousePointer = vbDefault
            Exit Function
        End If

        lCountLine = 0
        Do While Not lrsTab_Tipo_Fator.EOF
            lCountLine = lCountLine + 1
            lblnAchou = False
            For Each lobjstcFator In lobjstcPedLoc.ColecaoPedFator
                If lobjstcFator.Cod_Fator = lrsTab_Tipo_Fator("Cod_Fator") Then
                    lTemInfo = True
                    lblnAchou = True
                    lAuxiliar2 = Format(lCountLine, "00") & " (X) " & Trim(lrsTab_Tipo_Fator("Dsc_Fator"))
                    If Trim(lobjstcFator.Dsc_Compl_Fator) <> "" Then
                        lAuxiliar2 = lAuxiliar2 & " - " & Trim(lobjstcFator.Dsc_Compl_Fator)
                    End If
                    Print #pNumArqDados, lAuxiliar2
                    Exit For
                End If
            Next lobjstcFator
            If lblnAchou = False Then
                lAuxiliar2 = Format(lCountLine, "00") & " ( ) " & Trim(lrsTab_Tipo_Fator("Dsc_Fator"))
                Print #pNumArqDados, lAuxiliar2
            End If
            lrsTab_Tipo_Fator.MoveNext
        Loop
        Do While Not lCountLine > 24
            Print #pNumArqDados, " "
            lCountLine = lCountLine + 1
        Loop

        lCountLine = 0
        For Each lobjstcBem In lobjstcPedLoc.ColecaoPedLocBem
            lCountLine = lCountLine + 1
            lAuxiliar3 = lobjstcBem.gTipBem & " - " & lobjstcBem.Dsc_Bem
            lAuxiliar2 = Format(lCountLine, "000") & " - " & Left(lAuxiliar3, 100) & String(90 - Len(Left(lAuxiliar3, 90)), ".") & "R$ " & Format(lobjstcBem.Val_Bem, "standard")
            Print #pNumArqDados, lAuxiliar2
            lTemInfo = True
        Next lobjstcBem

        Do While lCountLine < 34
            lCountLine = lCountLine + 1
            Print #pNumArqDados, " "
        Loop

    Next lobjstcPedLoc

    Call gM46V999.gpFecha2(lbdTab_Tipo_Fator)

    mfCriDados_MULTCOTC = True

End Function
'1000308 - SyasMulti - Roubo - Cotação Coletiva com Planos Distintos nos Itens
Private Function mfCriDados_ROUBOCOT1(ByVal pNumArqDados As Integer, ByVal pobjstcPed As stcA46V702B, _
                                      pobjstcPedLoc As stcA46V708B, pNumDe As Integer, pNumAte As Integer) As Boolean
'Função: cria arquivo de dados - cotação ROUBO simples.

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
    Dim lPrazo As Integer
    Dim lCodCobert As Integer
    Dim lNumMeses As Integer
    Dim lDescCom As Double
    Dim lDescExp As Double
    Dim lValorComis As Single
    Dim lbdTbm_Regiao As ADODB.Connection
    Dim lrsTbm_Regiao As clsYasRecordSet
    Dim lQuaDias As Integer
    Dim lI2 As Long
    Dim lSelect As String

    mfCriDados_ROUBOCOT1 = False

    '001. Título do seguro (de acordo com a base de dados).
    '     Neste caso deve ser colocado YASUDA ROUBO.
    lAuxiliar2 = " YASUDA " & UCase(pobjstcPed.gDesc_Ramo)

    Do While Len(lAuxiliar2) < 36
        lAuxiliar2 = " " & lAuxiliar2 & " "
    Loop

    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    'Ficha 901223
    '001B. Código do Processo Susep.
    '1000308 - SyasMulti - Roubo - Cotação Coletiva com Planos Distintos nos Itens
    lAuxiliar2 = "Processo Susep 10.002505/01-16" & "              Pg " & pNumDe & "/" & pNumAte

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
    If pobjstcPed.ColecaoPedCorr.Count > 0 Then
        Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(1)
        lAuxiliar1 = lObjstcCorr.Cod_Corr
        lAuxiliar2 = lObjstcCorr.gNom_Corretor
    Else
        lAuxiliar1 = 0
        lAuxiliar2 = " "
    End If

    Print #pNumArqDados, Format$(lAuxiliar1, "0000") & " - " & lAuxiliar2

    '007. Tarifa .
    lAuxiliar2 = pobjstcPed.Cod_Produto
    If pobjstcPed.Cod_Sequencia = 1 Then
        lAuxiliar2 = lAuxiliar2 & " (A)"
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '008. Nosso número .
    lAuxiliar2 = pobjstcPed.NOSSO_NUMERO
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)


    If pobjstcPed.objstcCtrlEmis.Tip_Cotac = 2 And pobjstcPed.objstcCtrlEmis.Num_Cotac <> 0 Then
        Print #pNumArqDados, "Siscota: " & pobjstcPed.objstcCtrlEmis.Num_Cotac
    Else
        Print #pNumArqDados, " "
    End If
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
        'lAuxiliar2 = "CNPJ.........:"
        lAuxiliar2 = "              "
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        lAuxiliar2 = pobjstcPed.gCNPJ_CPF_Formatado & "   " & _
                     "Ramo de atividade: " & pobjstcPed.gDesc_RamoAtiv
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    Case e_TipPessoa_Fisica
        Print #pNumArqDados, "Física"
        '012 e 013. Proponente - CNPJ ou CPF .
        'lAuxiliar2 = "CPF...........:"
        lAuxiliar2 = "              "
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
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    lAuxiliar2 = " PLANO: " & pobjstcPed.Desc_Plano
    lAuxiliar2 = lAuxiliar2 & " - PRODUTO: " & Format(pobjstcPedLoc.Cod_Produto, "0000")
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '030. Dados do seguro - construção .
    'Ficha 1000305
    If gM46V999.gCod_Ramo = e_CodRamo_Empresarial Then
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
    If pobjstcPed.Cod_Produto <= gM46V999.TarifaAntiga Then
        lAuxiliar2 = pobjstcPedLoc.Cod_Regiao & " - " & Left(pobjstcPedLoc.gDesc_Regiao, 38)
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    Else
        'Nova Região
        If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTbm_Regiao, lMensagem) = False Then
            Call gM46V999.gpGraLog(2, lMensagem)
            Exit Function
        End If

        lSelect = "Select Cod_Faixa_Reg, DSC_REG from Tbm_Faixa_Regiao "
        lSelect = lSelect & " Where cod_Ramo = " & pobjstcPed.Cod_Ramo
        lSelect = lSelect & "    and Cod_Modalidade = " & pobjstcPedLoc.Cod_Produto
        lSelect = lSelect & "    and Cod_Tarif = " & pobjstcPed.Cod_Produto
        lSelect = lSelect & "    and COD_Sequencia <= " & pobjstcPed.Cod_Sequencia
        lSelect = lSelect & "    and Cep_Ini <= " & pobjstcPedLoc.Num_Cep_Risco
        lSelect = lSelect & "    and Cep_Fim >= " & pobjstcPedLoc.Num_Cep_Risco
        lSelect = lSelect & "    order by Cod_Sequencia desc "

        If gM46V999.gfObtRegistro(lbdTbm_Regiao, lSelect, lrsTbm_Regiao, lMensagem) = False Then
            Call gM46V999.gpGraLog(0, lMensagem)
            Call gM46V999.gpFecha2(lbdTbm_Regiao)
            Exit Function
        End If

        If Not lrsTbm_Regiao.EOF Then
            lAuxiliar2 = gM46V999.Check(lrsTbm_Regiao("Cod_Faixa_Reg")) & " - " & Left(gM46V999.Check(lrsTbm_Regiao("DSC_REG")), 38)
        Else
            lAuxiliar2 = "Não cadastrada."
        End If
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    End If

    '032. Dados do seguro - vigência .

    lAuxiliar2 = "de " & gM46V999.gfForData(2, pobjstcPed.Dat_Ini_Vig) & _
               " a " & gM46V999.gfForData(2, pobjstcPed.Dat_Ter_Vig)

    'Se for perfil Calculo não será impresso conforme ficha 0701831
    If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda And _
       (gM46V999.PerfilAcesso = e_Perfil_CALC Or gM46V999.PerfilAcesso = e_Perfil_CALC_Gestor) Then
        lQuaDias = gM46V999.gfDifDatas(pobjstcPed.Dat_Ini_Vig, pobjstcPed.Dat_Ter_Vig)
        If lQuaDias >= 365 Then
            lAuxiliar2 = ""
        End If
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '033. Dados do seguro - valor em risco .
    lAuxiliar2 = "R$ " & Format(pobjstcPedLoc.Val_Vr, "standard")
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '034. Dados do seguro - tipo de emissão .
    lAuxiliar2 = pobjstcPedLoc.gDesc_TipEmissao
    Select Case pobjstcPedLoc.Tip_Emissao
    Case e_TipEmissao_RenYasuda, e_TipEmissao_Endosso
        lAuxiliar2 = lAuxiliar2 & " " & "  Apólice: " & pobjstcPedLoc.Num_Apol_Anterior & "  Item: " & pobjstcPedLoc.Num_Item_Anterior
    End Select
    Print #pNumArqDados, lAuxiliar2

    Print #pNumArqDados, " "
    Print #pNumArqDados, " "

    'daqui prá baixo

    '037 a 212 Coberturas contratadas - linhas de 01 a 09.
    lCountLine = 0
    lValTotCober = 0
    For Each lobjstcCob In pobjstcPedLoc.ColecaoPedCobert
        lCountLine = lCountLine + 1
        'Código
        Print #pNumArqDados, Format$(lobjstcCob.Cod_Cobert, "0000")
        lCodCobert = lobjstcCob.Cod_Cobert
        'Descrição
        Print #pNumArqDados, Left(Trim(lobjstcCob.gNomeCobertura), 53)

        lAuxiliar1 = IIf(IsNumeric(lobjstcCob.Num_Meses_PI), lobjstcCob.Num_Meses_PI, 0)
        If lAuxiliar1 = 0 Then
            Print #pNumArqDados, " "

            Print #pNumArqDados, " "
        Else
            'pi mes
            Print #pNumArqDados, Format$(lobjstcCob.Num_Meses_PI, "00")
            lNumMeses = lobjstcCob.Num_Meses_PI
            Call mpPrazoTabCoefPI(pobjstcPed, lCodCobert, lNumMeses, lPrazo)
            'pi dias
            If lPrazo <> 0 Then
                Print #pNumArqDados, lPrazo
            Else
                Print #pNumArqDados, " "
            End If
        End If
        '%Sobre Prejuízos
        lAuxiliar1 = lobjstcCob.Per_Franq
        lAuxiliar2 = Format$(lAuxiliar1, "standard")
        lTamanho = Len(lAuxiliar2)
        If lTamanho < 10 Then
            lAuxiliar2 = Space(10 - lTamanho) & lAuxiliar2
        End If
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        'Limitado ao Mínimo
        lAuxiliar1 = lobjstcCob.Val_Franq_Min
        lAuxiliar2 = Format$(lAuxiliar1, "standard")
        lTamanho = Len(lAuxiliar2)
        If lTamanho < 14 Then
            lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
        End If
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

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
        lDescCom = lobjstcCob.Per_Desc_Coml
        lAuxiliar2 = Format$(lAuxiliar1, "standard")
        lTamanho = Len(lAuxiliar2)
        If lTamanho < 10 Then
            lAuxiliar2 = Space(10 - lTamanho) & lAuxiliar2
        End If

        lDescExp = 0

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

    Do While lCountLine < 9
        lCountLine = lCountLine + 1
        Print #pNumArqDados, " "
        Print #pNumArqDados, " "
        Print #pNumArqDados, " "
        Print #pNumArqDados, " "
        Print #pNumArqDados, " "
        Print #pNumArqDados, " "
        Print #pNumArqDados, " "
        Print #pNumArqDados, " "
    Loop

    Print #pNumArqDados, "              Soma:"
    '213. Coberturas contratadas - soma.
    lAuxiliar1 = lValTotCober
    lAuxiliar2 = Format$(lAuxiliar1, "standard")
    lTamanho = Len(lAuxiliar2)
    If lTamanho < 14 Then
        lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    'Calcular Soma da Comissão para saber se é premio NET
    If pobjstcPed.ColecaoPedCorr.Count > 0 Then
        Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(1)
        lValorComis = lObjstcCorr.Per_Corr
        If pobjstcPed.ColecaoPedCorr.Count > 1 Then
            Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(2)
            lValorComis = lValorComis + lObjstcCorr.Per_Corr
            If pobjstcPed.ColecaoPedCorr.Count > 2 Then
                Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(3)
                lValorComis = lValorComis + lObjstcCorr.Per_Corr
            End If
        End If
    End If

    If lValorComis = 0 Then
        Print #pNumArqDados, "        Prêmio NET:"
    Else
        Print #pNumArqDados, "            Prêmio:"
    End If

    '214. Coberturas contratadas - prêmio líquido.
    lAuxiliar1 = pobjstcPed.VAL_PR
    lAuxiliar2 = Format$(lAuxiliar1, "standard")
    lTamanho = Len(lAuxiliar2)
    If lTamanho < 14 Then
        lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '*************** Premio NET
    If lValorComis = 0 Then
        For lI = 1 To 6
            Print #pNumArqDados, " "
        Next
        GoTo ImprimirLegendas
    End If
    '*************** Premio NET
    Print #pNumArqDados, "  Custo da Apólice:"
    '215. Coberturas contratadas - custo da apólice.
    lAuxiliar1 = pobjstcPed.Val_Cust_Emis
    lAuxiliar2 = Format$(lAuxiliar1, "standard")
    lTamanho = Len(lAuxiliar2)
    If lTamanho < 14 Then
        lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    lAuxiliar2 = "IOF (" & Format(gM46V111.gProJurIOF, "standard") & "%):"
    Print #pNumArqDados, String(19 - Len(lAuxiliar2), " ") & lAuxiliar2

    '216. Coberturas contratadas - IOF.
    lAuxiliar1 = pobjstcPed.Val_IOF
    lAuxiliar2 = Format$(lAuxiliar1, "standard")
    lTamanho = Len(lAuxiliar2)
    If lTamanho < 14 Then
        lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    Print #pNumArqDados, "Prêmio total vista:"
    '217. Coberturas contratadas - prêmio total a vista.
    lAuxiliar1 = pobjstcPed.gValorTotal
    lAuxiliar2 = Format$(lAuxiliar1, "standard")
    lTamanho = Len(lAuxiliar2)
    If lTamanho < 14 Then
        lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    'Premio NET
ImprimirLegendas:

    '218. Informação cifrada .

    '   B: 25 / A: 50 / P: 10 / CO: 99999 / C: 99999 Onde:
    '   B = bônus
    '   A = aparelho
    '   P = progressivo
    '   CO = coeficiente de comissão
    '   C = coeficiente LMG/VR

    lAuxiliar2 = "B:" & Format(pobjstcPedLoc.Per_Bonus, "00")
    lAuxiliar2 = lAuxiliar2 & "/P:" & Format(pobjstcPedLoc.Per_Desc_Progres, "00")
    lAuxiliar1 = 0

    If pobjstcPed.ColecaoPedCorr.Count > 0 Then
        Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(1)
        lAuxiliar1 = lObjstcCorr.Per_Corr
        If pobjstcPed.ColecaoPedCorr.Count > 1 Then
            Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(2)
            lAuxiliar1 = lAuxiliar1 + lObjstcCorr.Per_Corr
            If pobjstcPed.ColecaoPedCorr.Count > 2 Then
                Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(3)
                lAuxiliar1 = lAuxiliar1 + lObjstcCorr.Per_Corr
            End If
        End If
    End If

    lAuxiliar2 = lAuxiliar2 & "/CO:" & Format(lAuxiliar1 * 100, "0000")
    lAuxiliar2 = lAuxiliar2 & "/DC:" & Format(lDescCom * 100, "0000")
    'lAuxiliar2 = lAuxiliar2 & "/PRGRS6"

    'Ficha 900550
    '/V1.0 -> versão após ajustar o bônus das coberturas
    lAuxiliar2 = lAuxiliar2 & "/V1.0"

    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)


    '*************** Premio NET
    If lValorComis = 0 Then
        For lI = 1 To 24
            Print #pNumArqDados, " "
        Next
        GoTo ImprimirImportante
    End If
    '*************** Premio NET


    'Calcula o fracionamento
    With pobjstcPed
        Call gM46V998.gfCalFracionamento(pobjstcPed, pobjstcPedLoc.Tip_Emissao, _
                                         lValorParc, lMensagem, gM46V111.gProJurIOF, .Cod_Reg, .Val_IOF, .Val_Parc_Min)
    End With
    '219 a 230. Formas de pagamento - primeira parcela a vista.

    For lI = 1 To 12
        If lI > gM46V999.gProQuaMaxParcelas Then
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
                    If lTamanho < 9 Then
                        lDemParcelas = Space(9 - lTamanho) & lDemParcelas
                    End If
                    'd) Juros.
                    Set lobjstcJur = New stcA46V711B
                    On Error Resume Next
                    Set lobjstcJur = pobjstcPed.ColecaoPedJuros.Item(pobjstcPed.NOSSO_NUMERO & Format(lI, "00"))
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

    '231 a 242. Formas de pagamento - primeira parcela a 30 dias.
    For lI = 1 To 12
        If lI > gM46V999.gProQuaMaxParcelas Then
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
                    If lTamanho < 10 Then
                        lPriParcela = Space(10 - lTamanho) & lPriParcela
                    End If
                    'c) Valor das demais parcelas.
                    lAuxiliar1 = lValorParc(lI, 4)
                    lDemParcelas = Format$(lAuxiliar1, "standard")
                    lTamanho = Len(lDemParcelas)
                    If lTamanho < 9 Then
                        lDemParcelas = Space(9 - lTamanho) & lDemParcelas
                    End If
                    'd) Juros.
                    Set lobjstcJur = New stcA46V711B
                    On Error Resume Next
                    Set lobjstcJur = pobjstcPed.ColecaoPedJuros.Item(pobjstcPed.NOSSO_NUMERO & Format(lI, "00"))
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
    'Premio NET
ImprimirImportante:

    '251 a 260. Importante.
    Call mpTextoImportante(pobjstcPedLoc, pobjstcPed, lDescImportante, "ROUBOCOT1", 36)
    lCountLine = 0
    For lI = 0 To UBound(lDescImportante) - 1
        If lDescImportante(lI) = "" Then Exit For
        lAuxiliar2 = lDescImportante(lI)
        If gM46V999.gfPreenchido(lAuxiliar2) Then
            lCountLine = lCountLine + 1
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        End If
    Next lI

    ReDim mDemaisObservacoes(36 - lCountLine)
    Call mpTextoObs(pobjstcPedLoc, lDescImportante, 36 - lCountLine)
    For lI = 0 To UBound(lDescImportante) - 1
        lAuxiliar2 = lDescImportante(lI)
        If lCountLine = 36 Then
            mDemaisObservacoes(lI2) = lAuxiliar2
            lI2 = lI2 + 1
        Else
            If gM46V999.gfPreenchido(lAuxiliar2) Then
                lCountLine = lCountLine + 1
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            End If
        End If
    Next lI
    Do While lCountLine < 36
        Print #pNumArqDados, " "
        lCountLine = lCountLine + 1
    Loop

    mfCriDados_ROUBOCOT1 = True
End Function


Private Function mfCriDados_ROUBOCOT2(ByVal pNumArqDados As Integer, ByVal pobjstcPed As stcA46V702B, _
                                      pobjstcPedLoc As stcA46V708B, pNumDe As Integer, _
                                      pNumAte As Integer) As Boolean

'Função: cria arquivo de dados - cotação ROUBO Anexo.

'Parâmetro de entrada...pNumArqDados...Número do arquivo de dados.


    Dim lAuxiliar1 As Double  'Auxiliar - utilizado apenas como exemplo.
    Dim lAuxiliar2 As String  'Auxiliar - utilizado apenas como exemplo.
    Dim lAuxiliar3 As String  'Auxiliar - utilizado apenas como exemplo.
    Dim lObjstcCorr As stcA46V705B    'Objeto stcA46V705B  (corretor).
    Dim lCountLine As Integer    'Contador de linhas.
    Dim lobjstcBem As stcA46V715B
    Dim lobjstcClau As stcA46V703B    'Objeto da classe clsA46V703A (TAB_PED_CLAU).
    Dim lMensagem As String
    Dim lbdTab_Tipo_Fator As ADODB.Connection   'Arquivo: P00Tarifa.
    Dim lrsTab_Tipo_Fator As clsYasRecordSet    'Registro :Tab_Tipo_Fator.
    Dim lobjstcFator As stcA46V720B   'Objeto da classe de estrutura stcA46V720B (TAB_PED_FAT_RISCO).
    Dim lblnAchou As Boolean    'Indica se achou  o fator ou não.
    Dim lQuant As Integer
    Dim lTotalPag As Integer
    Dim lTemInfo As Boolean
    Dim lValorBem As Double
    Dim lI2 As Long
    Dim lSelect As String

    mfCriDados_ROUBOCOT2 = False
    lTemInfo = False

    lQuant = 0
    If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTab_Tipo_Fator, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Function
    End If

    lQuant = 0
    lTotalPag = 0

    Call mfImpCabec_ROUBOCOT2(pNumArqDados, pobjstcPed, pobjstcPedLoc, pNumDe, pNumAte)

    'Antes de iniciar eu carrego as observações que não couberam na página anterior...

    lCountLine = 0
    For lQuant = 0 To UBound(mDemaisObservacoes) - 1
        If gM46V999.gfPreenchido(mDemaisObservacoes(lQuant)) Then
            Print #pNumArqDados, mDemaisObservacoes(lQuant)
            lCountLine = lCountLine + 1
        End If
    Next lQuant

    If lCountLine <> 0 Then
        Print #pNumArqDados, " "
        lCountLine = lCountLine + 1
    End If

    lQuant = 0

    '010 a 024. Fatores de Proteção - linhas de 01 a 25.

    'Ficha 1000601
    If gM46V999.gCod_Ramo = e_CodRamo_Residencial Then
        If pobjstcPedLoc.gLiberaDescontoEspecial = True Then
            lSelect = " Select distinct DSC_FATOR, COD_FATOR  from Tbm_Tipo_Fator_Protec "
            lSelect = lSelect & " Where cod_Ramo = " & pobjstcPed.Cod_Ramo
            lSelect = lSelect & " and Cod_Modalidade = " & pobjstcPedLoc.Cod_Produto
            lSelect = lSelect & " and COD_TARIF = " & pobjstcPed.Cod_Produto
            lSelect = lSelect & " and cod_Sequencia  <= " & pobjstcPed.Cod_Sequencia
            lSelect = lSelect & " order by DSC_FATOR"
        Else
            lSelect = " Select distinct DSC_FATOR, COD_FATOR  from Tbm_Tipo_Fator_Protec "
            lSelect = lSelect & " Where cod_Ramo = " & pobjstcPed.Cod_Ramo
            lSelect = lSelect & " and Cod_Modalidade = " & pobjstcPedLoc.Cod_Produto
            lSelect = lSelect & " and COD_TARIF = " & pobjstcPed.Cod_Produto
            lSelect = lSelect & " and cod_Sequencia  <= " & pobjstcPed.Cod_Sequencia
            lSelect = lSelect & " and Cod_Fator <> 7 "    'Não carregar o item Desconto Especial
            lSelect = lSelect & " order by DSC_FATOR"
        End If
    Else
        lSelect = " Select distinct DSC_FATOR, COD_FATOR  from Tbm_Tipo_Fator_Protec "
        lSelect = lSelect & " Where cod_Ramo = " & pobjstcPed.Cod_Ramo
        lSelect = lSelect & " and Cod_Modalidade = " & pobjstcPedLoc.Cod_Produto
        lSelect = lSelect & " and COD_TARIF = " & pobjstcPed.Cod_Produto
        lSelect = lSelect & " and cod_Sequencia  <= " & pobjstcPed.Cod_Sequencia
        lSelect = lSelect & " order by DSC_FATOR"
    End If

    If gM46V999.gfObtRegistro(lbdTab_Tipo_Fator, lSelect, lrsTab_Tipo_Fator, lMensagem) = False Then
        Call gM46V999.gpFecha2(lbdTab_Tipo_Fator)
        Call gM46V999.gpGraLog(2, lMensagem)
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Print #pNumArqDados, "### QUESTIONÁRIO ###"
    lCountLine = lCountLine + 1
    If lrsTab_Tipo_Fator.EOF Then
        Print #pNumArqDados, "Não informado"
        lCountLine = lCountLine + 1
        Print #pNumArqDados, " "
        lCountLine = lCountLine + 1
    End If


    Do While Not lrsTab_Tipo_Fator.EOF
        lblnAchou = False
        For Each lobjstcFator In pobjstcPedLoc.ColecaoPedFator
            If lobjstcFator.Cod_Fator = lrsTab_Tipo_Fator("Cod_Fator") Then
                lTemInfo = True
                lblnAchou = True
                lAuxiliar2 = Format(lCountLine, "00") & " (X) " & Trim(lrsTab_Tipo_Fator("Dsc_Fator"))
                If Trim(lobjstcFator.Dsc_Compl_Fator) <> "" Then
                    lAuxiliar2 = lAuxiliar2 & " - " & Trim(lobjstcFator.Dsc_Compl_Fator)
                End If
                Print #pNumArqDados, lAuxiliar2
                lCountLine = lCountLine + 1
                Exit For
            End If
        Next lobjstcFator
        If lblnAchou = False Then
            lAuxiliar2 = Format(lCountLine, "00") & " ( ) " & Trim(lrsTab_Tipo_Fator("Dsc_Fator"))
            Print #pNumArqDados, lAuxiliar2
            lCountLine = lCountLine + 1
        End If
        lrsTab_Tipo_Fator.MoveNext
    Loop
    Call gM46V999.gpFecha2(lbdTab_Tipo_Fator)

    lValorBem = 0

    Print #pNumArqDados, " "
    lCountLine = lCountLine + 1
    Print #pNumArqDados, "### RELAÇÃO DE BENS ###"
    lCountLine = lCountLine + 1
    If pobjstcPedLoc.ColecaoPedLocBem.Count = 0 Then
        Print #pNumArqDados, "Não informado"
        lCountLine = lCountLine + 1
        Print #pNumArqDados, " "
        lCountLine = lCountLine + 1
    End If


    lI2 = 0
    For Each lobjstcBem In pobjstcPedLoc.ColecaoPedLocBem
        If lCountLine > 84 Then
            lCountLine = 0
            lQuant = lQuant + 1
            lTotalPag = lTotalPag + 1
            Call mfImpCabec_ROUBOCOT2(pNumArqDados, pobjstcPed, pobjstcPedLoc, lQuant, lTotalPag)
            Print #pNumArqDados, " "
            lCountLine = lCountLine + 1
            Print #pNumArqDados, "### RELAÇÃO DE BENS ###"
            lCountLine = lCountLine + 1
        End If
        lAuxiliar3 = lobjstcBem.gTipBem & " - " & lobjstcBem.Dsc_Bem
        lI2 = lI2 + 1
        lAuxiliar2 = Format(lI2, "000") & " - " & Left(lAuxiliar3, 100) & String(90 - Len(Left(lAuxiliar3, 90)), ".") & "R$ " & Format(lobjstcBem.Val_Bem, "standard")
        lValorBem = lValorBem + lobjstcBem.Val_Bem
        Print #pNumArqDados, lAuxiliar2
        lCountLine = lCountLine + 1
        lTemInfo = True
    Next lobjstcBem

    If lValorBem <> 0 Then
        Print #pNumArqDados, " "
        lCountLine = lCountLine + 1
        Print #pNumArqDados, "Total: " & String(99 - Len(Left(lAuxiliar3, 99)), ".") & "R$ " & Format(lValorBem, "standard")
        lCountLine = lCountLine + 1
    End If


    'Clausulas
    Print #pNumArqDados, " "
    lCountLine = lCountLine + 1
    Print #pNumArqDados, "### CLÁUSULAS E CONDIÇÕES ###"
    lCountLine = lCountLine + 1
    If pobjstcPedLoc.ColecaoPedClau.Count = 0 Then
        Print #pNumArqDados, "Não informado"
        lCountLine = lCountLine + 1
        Print #pNumArqDados, " "
        lCountLine = lCountLine + 1
    End If

    For Each lobjstcClau In pobjstcPedLoc.ColecaoPedClau
        If lCountLine > 84 Then
            lCountLine = 0
            lQuant = lQuant + 1
            lTotalPag = lTotalPag + 1
            Call mfImpCabec_ROUBOCOT2(pNumArqDados, pobjstcPed, pobjstcPedLoc, lQuant, lTotalPag)
            Print #pNumArqDados, " "
            lCountLine = lCountLine + 1
            Print #pNumArqDados, "### CLÁUSULAS E CONDIÇÕES ###"
            lCountLine = lCountLine + 1
        End If
        lAuxiliar2 = lobjstcClau.gDescClausula
        Print #pNumArqDados, lAuxiliar2
        lCountLine = lCountLine + 1
        lTemInfo = True
    Next lobjstcClau

    Do While lCountLine < 84
        Print #pNumArqDados, " "
        lCountLine = lCountLine + 1
    Loop



    mfCriDados_ROUBOCOT2 = True

End Function


Private Function mfCriDados_ROUBOCOT3(ByVal pNumArqDados As Integer, ByVal pobjstcPed As stcA46V702B, _
                                      pobjstcPedLoc As stcA46V708B, pNumDe As Integer, _
                                      pNumAte As Integer) As Boolean
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
    Dim lObjstcCorr As stcA46V705B    'Objeto stcA46V705B  (corretor).
    Dim lobjstcJur As stcA46V711B
    Dim lCountLine As Integer
    Dim lValorParc As Variant
    Dim lValorComis As Single
    Dim lMensagem As String
    Dim lDescImportante() As String

    mfCriDados_ROUBOCOT3 = False

    '001. Título do documento.
    If pobjstcPed.Cod_Dizer = 0 Then
        Print #pNumArqDados, "COTAÇÃO DE SEGURO"
    Else
        Print #pNumArqDados, "COTAÇÃO DE ENDOSSO"
    End If

    '002. Título do seguro (de acordo com a base de dados).
    '     Neste caso deve ser colocado YASUDA ROUBO.
    Print #pNumArqDados, "YASUDA " & UCase(pobjstcPed.gDesc_Ramo)

    '003. Número da página.
    '1000308 - SyasMulti - Roubo - Cotação Coletiva com Planos Distintos nos Itens
    Print #pNumArqDados, "Página " & pNumDe & "/" & pNumAte

    '004. Emissão.
    Print #pNumArqDados, "Emissão: " & Format$(Now, "dd/mm/yyyy hh:nn:ss")

    'Susep
    lAuxiliar2 = "Processo Susep 10.002505/01-16"

    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

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
    If pobjstcPed.ColecaoPedCorr.Count > 0 Then
        Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(1)
        lAuxiliar1 = lObjstcCorr.Cod_Corr
        lAuxiliar2 = lObjstcCorr.gNom_Corretor
    Else
        lAuxiliar1 = 0
        lAuxiliar2 = " "
    End If

    'Soma da Comissão para saber se o prêmio é NET
    If pobjstcPed.ColecaoPedCorr.Count > 0 Then
        Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(1)
        lValorComis = lObjstcCorr.Per_Corr
        If pobjstcPed.ColecaoPedCorr.Count > 1 Then
            Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(2)
            lValorComis = lValorComis + lObjstcCorr.Per_Corr
            If pobjstcPed.ColecaoPedCorr.Count > 2 Then
                Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(3)
                lValorComis = lValorComis + lObjstcCorr.Per_Corr
            End If
        End If
    End If


    Print #pNumArqDados, Format$(lAuxiliar1, "0000") & " - " & lAuxiliar2

    '008. Tarifa .
    lAuxiliar2 = pobjstcPed.Cod_Produto
    If pobjstcPed.Cod_Sequencia = 1 Then
        lAuxiliar2 = lAuxiliar2 & " (A)"
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '009. Nosso número .
    lAuxiliar2 = pobjstcPed.NOSSO_NUMERO
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '    If pobjstcPed.objstcCtrlEmis.Tip_Cotac = 2 And pobjstcPed.objstcCtrlEmis.Num_Cotac <> 0 Then
    '        Print #pNumArqDados, "Siscota: " & pobjstcPed.objstcCtrlEmis.Num_Cotac
    '    Else
    '        Print #pNumArqDados, " "
    '    End If
    '
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
        'lAuxiliar2 = "CPF...........:"
        lAuxiliar2 = "               "
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        lAuxiliar2 = pobjstcPed.gCNPJ_CPF_Formatado & "   " & _
                     pobjstcPed.gCod_Doc_Segur & ": " & pobjstcPed.Num_Doc & "   " & _
                     "Órgão emissor: " & pobjstcPed.Org_Exp & "   " & _
                     "Emissão: " & gM46V999.gfForData(2, pobjstcPed.Dat_Emis_Doc)
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    Case e_TipPessoa_Juridica
        Print #pNumArqDados, "Jurídica"
        '013 e 014. Proponente - CNPJ ou CPF .
        'lAuxiliar2 = "CNPJ.........:"
        lAuxiliar2 = "              "
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
    If lValorComis = 0 Then
        Print #pNumArqDados, "          Prêmio NET:"
    Else
        Print #pNumArqDados, "Prêmio Líquido Total:"
    End If
    lAuxiliar1 = pobjstcPed.VAL_PR
    lAuxiliar2 = Format$(lAuxiliar1, "standard")
    lTamanho = Len(lAuxiliar2)
    If lTamanho < 14 Then
        lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '*******************************************************************************
    If lValorComis = 0 Then
        'Caso seja prêmio net só irá imprimir as observações
        For lI = 1 To 30
            Print #pNumArqDados, " "
        Next
        GoTo Finalizar
    End If

    '********************************************************************************
    Print #pNumArqDados, "    Custo da Apólice:"
    '023. Resumo: custo da apólice.
    lAuxiliar1 = pobjstcPed.Val_Cust_Emis
    lAuxiliar2 = Format$(lAuxiliar1, "standard")
    lTamanho = Len(lAuxiliar2)
    If lTamanho < 14 Then
        lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    lAuxiliar2 = "IOF (" & Format(gM46V111.gProJurIOF, "standard") & "%):"
    Print #pNumArqDados, String(21 - Len(lAuxiliar2), " ") & lAuxiliar2
    '024. Resumo: IOF.
    lAuxiliar1 = pobjstcPed.Val_IOF
    lAuxiliar2 = Format$(lAuxiliar1, "standard")
    lTamanho = Len(lAuxiliar2)
    If lTamanho < 14 Then
        lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)


    Print #pNumArqDados, "Prêmio total a vista:"
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
    For lI = 1 To 12
        If lI > gM46V999.gProQuaMaxParcelas Then
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
                Set lobjstcJur = pobjstcPed.ColecaoPedJuros.Item(pobjstcPed.NOSSO_NUMERO & Format(lI, "00"))
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
    For lI = 1 To 12
        If lI > gM46V999.gProQuaMaxParcelas Then
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
                Set lobjstcJur = pobjstcPed.ColecaoPedJuros.Item(pobjstcPed.NOSSO_NUMERO & Format(lI, "00"))
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

    'Caso seja prêmio net virá direto para cá
    '********************************************************************************

Finalizar:

    mfCriDados_ROUBOCOT3 = True
End Function


Private Function mfImpCabec_ROUBOCOT2(ByVal pNumArqDados As Integer, ByVal pobjstcPed As stcA46V702B, _
                                      pobjstcPedLoc As stcA46V708B, ByRef lQuant As Integer, ByRef lTotalPag As Integer) As Boolean

'Função: cria arquivo de dados - cotação ROUBO Anexo.

'Parâmetro de entrada...pNumArqDados...Número do arquivo de dados.


    Dim lAuxiliar1 As Double  'Auxiliar - utilizado apenas como exemplo.
    Dim lAuxiliar2 As String  'Auxiliar - utilizado apenas como exemplo.
    Dim lAuxiliar3 As String  'Auxiliar - utilizado apenas como exemplo.
    Dim lObjstcCorr As stcA46V705B    'Objeto stcA46V705B  (corretor).
    Dim lobjstcPedLoc As stcA46V708B    'Objeto stcA46V708B
    Dim lCountLine As Integer    'Contador de linhas.
    Dim lMensagem As String
    Dim lblnAchou As Boolean    'Indica se achou  o fator ou não.
    Dim lTemInfo As Boolean
    Dim lNumArquivo As Integer          'Número do arquivo.
    Dim lArqDados As String         'Endereço do arquivo de dados.

    mfImpCabec_ROUBOCOT2 = False
    lTemInfo = False

    '001. Título do seguro (de acordo com a base de dados).
    '     Neste caso deve ser colocado YASUDA ROUBO
    lAuxiliar2 = "YASUDA " & UCase(pobjstcPed.gDesc_Ramo)

    Do While Len(lAuxiliar2) < 45
        lAuxiliar2 = " " & lAuxiliar2 & " "
    Loop

    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '002. Página.
    Print #pNumArqDados, "Página: " & lQuant & "/" & lTotalPag

    '003. Emissão.
    Print #pNumArqDados, "Emissão: " & Format$(Now, "dd/mm/yyyy hh:nn:ss")

    'Ficha 901223
    '002B. Código do Processo Susep.
    lAuxiliar2 = "Processo Susep 10.002505/01-16"

    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)


    '004. Validade.
    If pobjstcPedLoc.Tip_Emissao = e_TipEmissao_RenYasuda Then
        Print #pNumArqDados, "Validade: " & Format$(CDate(gM46V999.gfForData(2, pobjstcPed.Dat_Pri_Calculo)) + 30, "dd/mm/yyyy")
    Else
        Print #pNumArqDados, "Validade: " & Format$(CDate(gM46V999.gfForData(2, pobjstcPed.Dat_Pri_Calculo)) + 15, "dd/mm/yyyy")
    End If

    '005. Número do pedido .
    lAuxiliar1 = pobjstcPed.Num_Ped
    Print #pNumArqDados, Format$(lAuxiliar1, "000000")

    '006. Número do item .
    lAuxiliar1 = pobjstcPedLoc.Num_Item
    Print #pNumArqDados, Format$(lAuxiliar1, "0000")

    '007. Corretor (código e nome) .
    If pobjstcPed.ColecaoPedCorr.Count > 0 Then
        Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(1)
        lAuxiliar1 = lObjstcCorr.Cod_Corr
        lAuxiliar2 = lObjstcCorr.gNom_Corretor
    Else
        lAuxiliar1 = 0
        lAuxiliar2 = " "
    End If
    Print #pNumArqDados, Format$(lAuxiliar1, "0000") & " - " & lAuxiliar2

    '008. Tarifa .
    lAuxiliar2 = pobjstcPed.Cod_Produto
    If pobjstcPed.Cod_Sequencia = 1 Then
        lAuxiliar2 = lAuxiliar2 & " (A)"
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '009. Nosso número .
    lAuxiliar2 = pobjstcPed.NOSSO_NUMERO
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    mfImpCabec_ROUBOCOT2 = True

End Function
Private Function mfCriDados_ROUBOCOT4(ByVal pNumArqDados As Integer, ByVal pobjstcPed As stcA46V702B, _
                                      pobjstcPedLoc As stcA46V708B, pNumDe As Integer, _
                                      pNumAte As Integer) As Boolean

'Função: cria arquivo de dados - cotação coletiva - dados do item -

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
    Dim lPrazo As Integer
    Dim lCodCobert As Integer
    Dim lNumMeses As Integer
    Dim lDescCom As Double
    Dim lDescExp As Double
    Dim lValorComis As Single
    Dim lbdTbm_Regiao As ADODB.Connection
    Dim lrsTbm_Regiao As clsYasRecordSet
    Dim lQuaDias As Integer
    Dim lobjstcPedLoc As stcA46V708B
    Dim lCountItem As Long
    Dim lI2 As Long
    Dim lSelect As String

    mfCriDados_ROUBOCOT4 = False

    ReDim mDemaisObservacoes(pobjstcPed.ColecaoPedLoc.Count, 36)
    lCountItem = 0
    For Each lobjstcPedLoc In pobjstcPed.ColecaoPedLoc
        '001. Título do seguro (de acordo com a base de dados).
        '     Neste caso deve ser colocado YASUDA ROUBO.
        lCountItem = lCountItem + 1
        lAuxiliar2 = " YASUDA " & UCase(pobjstcPed.gDesc_Ramo)

        Do While Len(lAuxiliar2) < 36
            lAuxiliar2 = " " & lAuxiliar2 & " "
        Loop

        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '002. Emissão.
        Print #pNumArqDados, "Emissão: " & Format$(Now, "dd/mm/yyyy hh:nn:ss")

        'Susep
        '1000308 - SyasMulti - Roubo - Cotação Coletiva com Planos Distintos nos Itens
        lAuxiliar2 = "Processo Susep 10.002505/01-16" & "             Pg " & pNumDe + Int(lobjstcPedLoc.Num_Item) - 1 & "/" & pNumAte

        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)


        '003. Validade.
        If lobjstcPedLoc.Tip_Emissao = e_TipEmissao_RenYasuda Then
            Print #pNumArqDados, "Validade: " & Format$(CDate(gM46V999.gfForData(2, pobjstcPed.Dat_Pri_Calculo)) + 30, "dd/mm/yyyy")
        Else
            Print #pNumArqDados, "Validade: " & Format$(CDate(gM46V999.gfForData(2, pobjstcPed.Dat_Pri_Calculo)) + 15, "dd/mm/yyyy")
        End If

        '004. Número do pedido .
        lAuxiliar1 = pobjstcPed.Num_Ped
        Print #pNumArqDados, Format$(lAuxiliar1, "000000")

        '005. Número do item .
        lAuxiliar1 = lobjstcPedLoc.Num_Item
        Print #pNumArqDados, Format$(lAuxiliar1, "0000")

        '006. Corretor (código e nome) .
        If pobjstcPed.ColecaoPedCorr.Count > 0 Then
            Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(1)
            lAuxiliar1 = lObjstcCorr.Cod_Corr
            lAuxiliar2 = lObjstcCorr.gNom_Corretor
        Else
            lAuxiliar1 = 0
            lAuxiliar2 = " "
        End If

        Print #pNumArqDados, Format$(lAuxiliar1, "0000") & " - " & lAuxiliar2

        '007. Tarifa .
        lAuxiliar2 = pobjstcPed.Cod_Produto
        If pobjstcPed.Cod_Sequencia = 1 Then
            lAuxiliar2 = lAuxiliar2 & " (A)"
        End If
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '008. Nosso número .
        lAuxiliar2 = pobjstcPed.NOSSO_NUMERO
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)


        If pobjstcPed.objstcCtrlEmis.Tip_Cotac = 2 And pobjstcPed.objstcCtrlEmis.Num_Cotac <> 0 Then
            Print #pNumArqDados, "Siscota: " & pobjstcPed.objstcCtrlEmis.Num_Cotac
        Else
            Print #pNumArqDados, " "
        End If
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
            'lAuxiliar2 = "CNPJ.........:"
            lAuxiliar2 = "              "
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            lAuxiliar2 = pobjstcPed.gCNPJ_CPF_Formatado & "   " & _
                         "Ramo de atividade: " & pobjstcPed.gDesc_RamoAtiv
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        Case e_TipPessoa_Fisica
            Print #pNumArqDados, "Física"
            '012 e 013. Proponente - CNPJ ou CPF .
            'lAuxiliar2 = "CPF...........:"
            lAuxiliar2 = "              "
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
        lAuxiliar2 = lobjstcPedLoc.Cod_Tip_Logr_Risco & " " & lobjstcPedLoc.Nom_Logr_Risco
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '021. Local em risco - número .
        lAuxiliar2 = lobjstcPedLoc.objstcPedInsp.Num_Logr_Risco
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '022. Local em risco - complemento .
        lAuxiliar2 = lobjstcPedLoc.Dsc_Compl_Risco
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '023. Local em risco - bairro .
        lAuxiliar2 = lobjstcPedLoc.Nom_Bairro_Risco
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '024. Local em risco - cidade .
        lAuxiliar2 = lobjstcPedLoc.Nom_Cid_Risco
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '025. Local em risco - UF .
        lAuxiliar2 = lobjstcPedLoc.Sig_UF_Risco
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '026. Local em risco - CEP .
        lAuxiliar2 = lobjstcPedLoc.gCEP_Formatado
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '027. Dados para inspeção - nome da pessoa para contato .
        lAuxiliar2 = lobjstcPedLoc.objstcPedInsp.Pess_Cont
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '028. Dados para inspeção - telefone para contato .
        lAuxiliar2 = lobjstcPedLoc.objstcPedInsp.Num_Tel
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '029. Dados do seguro - classificação .
        lAuxiliar2 = lobjstcPedLoc.Cod_Clasf & "-" & lobjstcPedLoc.Cod_Compl_Clasf & "- " & lobjstcPedLoc.Dsc_Clasf
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        'Ficha 1000308 - Permitir itens com planos diferentes no produto Roubo (cálculo coletivo)
        Call gM46V111.gpObtModalidade(lobjstcPedLoc.Tab_Clasf, lobjstcPedLoc.Cod_Ramo)
        lAuxiliar2 = " PLANO: " & gM46V111.gModalidade

        lAuxiliar2 = lAuxiliar2 & " - PRODUTO: " & Format(lobjstcPedLoc.Cod_Produto, "0000")
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '030. Dados do seguro - construção .
        'Ficha 1000305
        If gM46V999.gCod_Ramo = e_CodRamo_Empresarial Then
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

        '031. Dados do seguro - região .
        If pobjstcPed.Cod_Produto <= gM46V999.TarifaAntiga Then
            lAuxiliar2 = lobjstcPedLoc.Cod_Regiao & " - " & Left(lobjstcPedLoc.gDesc_Regiao, 38)
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        Else
            'Nova Região
            If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTbm_Regiao, lMensagem) = False Then
                Call gM46V999.gpGraLog(2, lMensagem)
                Exit Function
            End If

            lSelect = "Select Cod_Faixa_Reg, DSC_REG from Tbm_Faixa_Regiao "
            lSelect = lSelect & " Where cod_Ramo = " & pobjstcPed.Cod_Ramo
            lSelect = lSelect & "    and Cod_Modalidade = " & lobjstcPedLoc.Cod_Produto
            lSelect = lSelect & "    and Cod_Tarif = " & pobjstcPed.Cod_Produto
            lSelect = lSelect & "    and COD_Sequencia <= " & pobjstcPed.Cod_Sequencia
            lSelect = lSelect & "    and Cep_Ini <= " & lobjstcPedLoc.Num_Cep_Risco
            lSelect = lSelect & "    and Cep_Fim >= " & lobjstcPedLoc.Num_Cep_Risco
            lSelect = lSelect & "    order by Cod_Sequencia desc "

            If gM46V999.gfObtRegistro(lbdTbm_Regiao, lSelect, lrsTbm_Regiao, lMensagem) = False Then
                Call gM46V999.gpGraLog(0, lMensagem)
                Call gM46V999.gpFecha2(lbdTbm_Regiao)
                Exit Function
            End If

            If Not lrsTbm_Regiao.EOF Then
                lAuxiliar2 = gM46V999.Check(lrsTbm_Regiao("Cod_Faixa_Reg")) & " - " & Left(gM46V999.Check(lrsTbm_Regiao("DSC_REG")), 38)
            Else
                lAuxiliar2 = "Não cadastrada."
            End If
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        End If

        '032. Dados do seguro - vigência .

        lAuxiliar2 = "de " & gM46V999.gfForData(2, pobjstcPed.Dat_Ini_Vig) & _
                   " a " & gM46V999.gfForData(2, pobjstcPed.Dat_Ter_Vig)

        'Se for perfil Calculo não será impresso conforme ficha 0701831
        If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda And _
           (gM46V999.PerfilAcesso = e_Perfil_CALC Or gM46V999.PerfilAcesso = e_Perfil_CALC_Gestor) Then
            lQuaDias = gM46V999.gfDifDatas(pobjstcPed.Dat_Ini_Vig, pobjstcPed.Dat_Ter_Vig)
            If lQuaDias >= 365 Then
                lAuxiliar2 = ""
            End If
        End If
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '033. Dados do seguro - valor em risco .
        lAuxiliar2 = "R$ " & Format(lobjstcPedLoc.Val_Vr, "standard")
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '034. Dados do seguro - tipo de emissão .
        lAuxiliar2 = lobjstcPedLoc.gDesc_TipEmissao
        Select Case lobjstcPedLoc.Tip_Emissao
        Case e_TipEmissao_RenYasuda, e_TipEmissao_Endosso
            lAuxiliar2 = lAuxiliar2 & " " & "  Apólice: " & lobjstcPedLoc.Num_Apol_Anterior & "  Item: " & lobjstcPedLoc.Num_Item_Anterior
        End Select
        Print #pNumArqDados, lAuxiliar2

        Print #pNumArqDados, " "
        Print #pNumArqDados, " "

        'daqui prá baixo

        '037 a 212 Coberturas contratadas - linhas de 01 a 09.
        lCountLine = 0
        lValTotCober = 0
        For Each lobjstcCob In lobjstcPedLoc.ColecaoPedCobert
            lCountLine = lCountLine + 1
            'Código
            Print #pNumArqDados, Format$(lobjstcCob.Cod_Cobert, "0000")
            lCodCobert = lobjstcCob.Cod_Cobert
            'Descrição
            Print #pNumArqDados, Left(lobjstcCob.gNomeCobertura, 53)

            lAuxiliar1 = IIf(IsNumeric(lobjstcCob.Num_Meses_PI), lobjstcCob.Num_Meses_PI, 0)
            If lAuxiliar1 = 0 Then
                Print #pNumArqDados, " "

                Print #pNumArqDados, " "
            Else
                'pi mes
                Print #pNumArqDados, Format$(lobjstcCob.Num_Meses_PI, "00")
                lNumMeses = lobjstcCob.Num_Meses_PI
                Call mpPrazoTabCoefPI(pobjstcPed, lCodCobert, lNumMeses, lPrazo)
                'pi dias
                If lPrazo <> 0 Then
                    Print #pNumArqDados, lPrazo
                Else
                    Print #pNumArqDados, " "
                End If
            End If
            '%Sobre Prejuízos
            lAuxiliar1 = lobjstcCob.Per_Franq
            lAuxiliar2 = Format$(lAuxiliar1, "standard")
            lTamanho = Len(lAuxiliar2)
            If lTamanho < 10 Then
                lAuxiliar2 = Space(10 - lTamanho) & lAuxiliar2
            End If
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

            'Limitado ao Mínimo
            lAuxiliar1 = lobjstcCob.Val_Franq_Min
            lAuxiliar2 = Format$(lAuxiliar1, "standard")
            lTamanho = Len(lAuxiliar2)
            If lTamanho < 14 Then
                lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
            End If
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

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
            lDescCom = lobjstcCob.Per_Desc_Coml
            lAuxiliar2 = Format$(lAuxiliar1, "standard")
            lTamanho = Len(lAuxiliar2)
            If lTamanho < 10 Then
                lAuxiliar2 = Space(10 - lTamanho) & lAuxiliar2
            End If

            lDescExp = 0

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

        Do While lCountLine < 9
            lCountLine = lCountLine + 1
            Print #pNumArqDados, " "
            Print #pNumArqDados, " "
            Print #pNumArqDados, " "
            Print #pNumArqDados, " "
            Print #pNumArqDados, " "
            Print #pNumArqDados, " "
            Print #pNumArqDados, " "
            Print #pNumArqDados, " "
        Loop

        Print #pNumArqDados, "              Soma:"
        '213. Coberturas contratadas - soma.
        lAuxiliar1 = lValTotCober
        lAuxiliar2 = Format$(lAuxiliar1, "standard")
        lTamanho = Len(lAuxiliar2)
        If lTamanho < 14 Then
            lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
        End If
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        'Calcular Soma da Comissão para saber se é premio NET
        If pobjstcPed.ColecaoPedCorr.Count > 0 Then
            Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(1)
            lValorComis = lObjstcCorr.Per_Corr
            If pobjstcPed.ColecaoPedCorr.Count > 1 Then
                Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(2)
                lValorComis = lValorComis + lObjstcCorr.Per_Corr
                If pobjstcPed.ColecaoPedCorr.Count > 2 Then
                    Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(3)
                    lValorComis = lValorComis + lObjstcCorr.Per_Corr
                End If
            End If
        End If

        If lValorComis = 0 Then
            Print #pNumArqDados, "        Prêmio NET:"
        Else
            Print #pNumArqDados, "            Prêmio:"
        End If

        '214. Coberturas contratadas - prêmio líquido.
        lAuxiliar1 = pobjstcPed.VAL_PR
        lAuxiliar2 = Format$(lAuxiliar1, "standard")
        lTamanho = Len(lAuxiliar2)
        If lTamanho < 14 Then
            lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
        End If
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        '*************** Premio NET
        If lValorComis = 0 Then
            For lI = 1 To 6
                Print #pNumArqDados, " "
            Next
            GoTo ImprimirLegendas
        End If
        '*************** Premio NET
        Print #pNumArqDados, "  Custo da Apólice:"
        '215. Coberturas contratadas - custo da apólice.
        lAuxiliar1 = pobjstcPed.Val_Cust_Emis
        lAuxiliar2 = Format$(lAuxiliar1, "standard")
        lTamanho = Len(lAuxiliar2)
        If lTamanho < 14 Then
            lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
        End If
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        lAuxiliar2 = "IOF (" & Format(gM46V111.gProJurIOF, "standard") & "%):"
        Print #pNumArqDados, String(19 - Len(lAuxiliar2), " ") & lAuxiliar2

        '216. Coberturas contratadas - IOF.
        lAuxiliar1 = pobjstcPed.Val_IOF
        lAuxiliar2 = Format$(lAuxiliar1, "standard")
        lTamanho = Len(lAuxiliar2)
        If lTamanho < 14 Then
            lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
        End If
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        Print #pNumArqDados, "Prêmio total vista:"
        '217. Coberturas contratadas - prêmio total a vista.
        lAuxiliar1 = pobjstcPed.gValorTotal
        lAuxiliar2 = Format$(lAuxiliar1, "standard")
        lTamanho = Len(lAuxiliar2)
        If lTamanho < 14 Then
            lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
        End If
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

        'Premio NET
ImprimirLegendas:

        '218. Informação cifrada .

        '   B: 25 / A: 50 / P: 10 / CO: 99999 / C: 99999 Onde:
        '   B = bônus
        '   A = aparelho
        '   P = progressivo
        '   CO = coeficiente de comissão
        '   C = coeficiente LMG/VR

        lAuxiliar2 = "B:" & Format(lobjstcPedLoc.Per_Bonus, "00")
        lAuxiliar2 = lAuxiliar2 & "/P:" & Format(lobjstcPedLoc.Per_Desc_Progres, "00")
        lAuxiliar1 = 0

        If pobjstcPed.ColecaoPedCorr.Count > 0 Then
            Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(1)
            lAuxiliar1 = lObjstcCorr.Per_Corr
            If pobjstcPed.ColecaoPedCorr.Count > 1 Then
                Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(2)
                lAuxiliar1 = lAuxiliar1 + lObjstcCorr.Per_Corr
                If pobjstcPed.ColecaoPedCorr.Count > 2 Then
                    Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(3)
                    lAuxiliar1 = lAuxiliar1 + lObjstcCorr.Per_Corr
                End If
            End If
        End If

        lAuxiliar2 = lAuxiliar2 & "/CO:" & Format(lAuxiliar1 * 100, "0000")
        lAuxiliar2 = lAuxiliar2 & "/DC:" & Format(lDescCom * 100, "0000")
        'lAuxiliar2 = lAuxiliar2 & "/PRGRS6"

        'Ficha 900550
        '/V1.0 -> versão após ajustar o bônus das coberturas
        lAuxiliar2 = lAuxiliar2 & "/V1.0"

        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)


        '*************** Premio NET
        'If lValorComis = 0 Then
        '    For lI = 1 To 24
        '        Print #pNumArqDados, " "
        '    Next
        '    GoTo ImprimirImportante
        'End If
        '*************** Premio NET


        'Calcula o fracionamento

        'Premio NET
ImprimirImportante:

        '251 a 260. Importante.
        Call mpTextoImportante(lobjstcPedLoc, pobjstcPed, lDescImportante, "ROUBOCOT4", 36)
        lCountLine = 0
        For lI = 0 To UBound(lDescImportante) - 1
            lAuxiliar2 = lDescImportante(lI)
            'If lCountLine = 36 Then Exit For
            If lDescImportante(lI) = "" Then Exit For
            If gM46V999.gfPreenchido(lAuxiliar2) Then
                lCountLine = lCountLine + 1
                Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            End If
        Next lI
        Call mpTextoObs(lobjstcPedLoc, lDescImportante, 36 - lCountLine)
        lI2 = 0
        For lI = 0 To UBound(lDescImportante) - 1
            lAuxiliar2 = lDescImportante(lI)
            If lDescImportante(lI) = "" Then Exit For
            If lCountLine = 36 Then
                mDemaisObservacoes(lCountItem, lI2) = lAuxiliar2
                lI2 = lI2 + 1
            Else
                If gM46V999.gfPreenchido(lAuxiliar2) Then
                    lCountLine = lCountLine + 1
                    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
                End If
            End If
        Next lI
        Do While lCountLine < 37
            Print #pNumArqDados, " "
            lCountLine = lCountLine + 1
        Loop
    Next lobjstcPedLoc

    mfCriDados_ROUBOCOT4 = True

End Function

Private Function mfCriDados_ROUBOCOT5(ByVal pNumArqDados As Integer, ByVal pobjstcPed As stcA46V702B, _
                                      pobjstcPedLoc As stcA46V708B, pNumDe As Integer, _
                                      pNumAte As Integer) As Boolean

'Função: cria arquivo de dados - cotação ROUBO Anexo.

'Parâmetro de entrada...pNumArqDados...Número do arquivo de dados.


    Dim lAuxiliar1 As Double  'Auxiliar - utilizado apenas como exemplo.
    Dim lAuxiliar2 As String  'Auxiliar - utilizado apenas como exemplo.
    Dim lAuxiliar3 As String  'Auxiliar - utilizado apenas como exemplo.
    Dim lObjstcCorr As stcA46V705B    'Objeto stcA46V705B  (corretor).
    Dim lCountLine As Integer    'Contador de linhas.
    Dim lobjstcBem As stcA46V715B
    Dim lobjstcClau As stcA46V703B    'Objeto da classe clsA46V703A (TAB_PED_CLAU).
    Dim lMensagem As String
    Dim lbdTab_Tipo_Fator As ADODB.Connection   'Arquivo: P00Tarifa.
    Dim lrsTab_Tipo_Fator As clsYasRecordSet    'Registro :Tab_Tipo_Fator.
    Dim lobjstcFator As stcA46V720B   'Objeto da classe de estrutura stcA46V720B (TAB_PED_FAT_RISCO).
    Dim lblnAchou As Boolean    'Indica se achou  o fator ou não.
    Dim lQuant As Integer
    Dim lTotalPag As Integer
    Dim lTemInfo As Boolean
    Dim lValorBem As Double
    Dim lobjstcPedLoc As stcA46V708B
    Dim lI2 As Long
    Dim lCountItem As Long
    Dim lSelect As String

    mfCriDados_ROUBOCOT5 = False
    lTemInfo = False
    If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTab_Tipo_Fator, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Function
    End If

    For Each lobjstcPedLoc In pobjstcPed.ColecaoPedLoc
        lCountLine = 0
        lCountItem = lCountItem + 1
        '1000308 - SyasMulti - Roubo - Cotação Coletiva com Planos Distintos nos Itens
        Call mfImpCabec_ROUBOCOT2(pNumArqDados, pobjstcPed, lobjstcPedLoc, pNumDe + Int(lobjstcPedLoc.Num_Item) - 1, pNumAte)

        For lI2 = 0 To 36
            If gM46V999.gfPreenchido(mDemaisObservacoes(lCountItem, lI2)) Then
                Print #pNumArqDados, mDemaisObservacoes(lCountItem, lI2)
                lCountLine = lCountLine + 1
            End If

        Next lI2
        If lCountLine <> 0 Then
            Print #pNumArqDados, " "
            lCountLine = lCountLine + 1
        End If
        '010 a 024. Fatores de Proteção - linhas de 01 a 25.

        'Ficha 1000601
        If gM46V999.gCod_Ramo = e_CodRamo_Residencial Then
            If lobjstcPedLoc.gLiberaDescontoEspecial = True Then
                lSelect = " Select distinct DSC_FATOR, COD_FATOR  from Tbm_Tipo_Fator_Protec "
                lSelect = lSelect & " Where cod_Ramo = " & pobjstcPed.Cod_Ramo
                lSelect = lSelect & " and Cod_Modalidade = " & lobjstcPedLoc.Cod_Produto
                lSelect = lSelect & " and COD_TARIF = " & pobjstcPed.Cod_Produto
                lSelect = lSelect & " and cod_Sequencia  <= " & pobjstcPed.Cod_Sequencia
                lSelect = lSelect & " order by DSC_FATOR"
            Else
                lSelect = " Select distinct DSC_FATOR, COD_FATOR  from Tbm_Tipo_Fator_Protec "
                lSelect = lSelect & " Where cod_Ramo = " & pobjstcPed.Cod_Ramo
                lSelect = lSelect & " and Cod_Modalidade = " & lobjstcPedLoc.Cod_Produto
                lSelect = lSelect & " and COD_TARIF = " & pobjstcPed.Cod_Produto
                lSelect = lSelect & " and cod_Sequencia  <= " & pobjstcPed.Cod_Sequencia
                lSelect = lSelect & " and Cod_Fator <> 7 "    'Não carregar o item Desconto Especial
                lSelect = lSelect & " order by DSC_FATOR"
            End If
        Else
            lSelect = " Select distinct DSC_FATOR, COD_FATOR  from Tbm_Tipo_Fator_Protec "
            lSelect = lSelect & " Where cod_Ramo = " & pobjstcPed.Cod_Ramo
            lSelect = lSelect & " and Cod_Modalidade = " & lobjstcPedLoc.Cod_Produto
            lSelect = lSelect & " and COD_TARIF = " & pobjstcPed.Cod_Produto
            lSelect = lSelect & " and cod_Sequencia  <= " & pobjstcPed.Cod_Sequencia
            lSelect = lSelect & " order by DSC_FATOR"
        End If

        If gM46V999.gfObtRegistro(lbdTab_Tipo_Fator, lSelect, lrsTab_Tipo_Fator, lMensagem) = False Then
            Call gM46V999.gpFecha2(lbdTab_Tipo_Fator)
            Call gM46V999.gpGraLog(2, lMensagem)
            Screen.MousePointer = vbDefault
            Exit Function
        End If

        Print #pNumArqDados, "### QUESTIONÁRIO ###"
        lCountLine = lCountLine + 1
        If lrsTab_Tipo_Fator.EOF Then
            Print #pNumArqDados, "Não informado"
            lCountLine = lCountLine + 1
            Print #pNumArqDados, " "
            lCountLine = lCountLine + 1
        End If

        Do While Not lrsTab_Tipo_Fator.EOF
            lblnAchou = False
            For Each lobjstcFator In lobjstcPedLoc.ColecaoPedFator
                If lobjstcFator.Cod_Fator = lrsTab_Tipo_Fator("Cod_Fator") Then
                    lTemInfo = True
                    lblnAchou = True
                    lAuxiliar2 = Format(lCountLine, "00") & " (X) " & Trim(lrsTab_Tipo_Fator("Dsc_Fator"))
                    If Trim(lobjstcFator.Dsc_Compl_Fator) <> "" Then
                        lAuxiliar2 = lAuxiliar2 & " - " & Trim(lobjstcFator.Dsc_Compl_Fator)
                    End If
                    Print #pNumArqDados, lAuxiliar2
                    lCountLine = lCountLine + 1
                    Exit For
                End If
            Next lobjstcFator
            If lblnAchou = False Then
                lAuxiliar2 = Format(lCountLine, "00") & " ( ) " & Trim(lrsTab_Tipo_Fator("Dsc_Fator"))
                Print #pNumArqDados, lAuxiliar2
                lCountLine = lCountLine + 1
            End If
            lrsTab_Tipo_Fator.MoveNext
        Loop

        Print #pNumArqDados, " "
        lCountLine = lCountLine + 1
        Print #pNumArqDados, "### RELAÇÃO DE BENS ###"
        lCountLine = lCountLine + 1
        If lobjstcPedLoc.ColecaoPedLocBem.Count = 0 Then
            Print #pNumArqDados, "Não informado"
            lCountLine = lCountLine + 1
            Print #pNumArqDados, " "
            lCountLine = lCountLine + 1
        End If
        lValorBem = 0
        lI2 = 0
        For Each lobjstcBem In lobjstcPedLoc.ColecaoPedLocBem
            If lCountLine > 84 Then
                lCountLine = 0
                lQuant = lQuant + 1
                lTotalPag = lTotalPag + 1
                '1000308 - SyasMulti - Roubo - Cotação Coletiva com Planos Distintos nos Itens
                Call mfImpCabec_ROUBOCOT2(pNumArqDados, pobjstcPed, lobjstcPedLoc, pNumDe + Int(lobjstcPedLoc.Num_Item) - 1, pNumAte)
                Print #pNumArqDados, " "
                lCountLine = lCountLine + 1
                Print #pNumArqDados, "### RELAÇÃO DE BENS ###"
                lCountLine = lCountLine + 1
            End If
            lAuxiliar3 = lobjstcBem.gTipBem & " - " & lobjstcBem.Dsc_Bem
            lI2 = lI2 + 1
            lAuxiliar2 = Format(lI2, "000") & " - " & Left(lAuxiliar3, 100) & String(90 - Len(Left(lAuxiliar3, 90)), ".") & "R$ " & Format(lobjstcBem.Val_Bem, "standard")
            lValorBem = lValorBem + lobjstcBem.Val_Bem
            Print #pNumArqDados, lAuxiliar2
            lCountLine = lCountLine + 1
            lTemInfo = True
        Next lobjstcBem

        If lValorBem <> 0 Then
            Print #pNumArqDados, " "
            lCountLine = lCountLine + 1
            Print #pNumArqDados, "Total: " & String(97 - Len(Left(lAuxiliar3, 97)), ".") & "R$ " & Format(lValorBem, "standard")
            lCountLine = lCountLine + 1
        End If


        'Clausulas
        Print #pNumArqDados, " "
        lCountLine = lCountLine + 1
        Print #pNumArqDados, "### CLÁUSULAS E CONDIÇÕES ###"
        lCountLine = lCountLine + 1
        If lobjstcPedLoc.ColecaoPedClau.Count = 0 Then
            Print #pNumArqDados, "Não informado"
            lCountLine = lCountLine + 1
            Print #pNumArqDados, " "
            lCountLine = lCountLine + 1
        End If

        For Each lobjstcClau In lobjstcPedLoc.ColecaoPedClau
            If lCountLine > 84 Then
                lCountLine = 0
                lQuant = lQuant + 1
                lTotalPag = lTotalPag + 1
                '1000308 - SyasMulti - Roubo - Cotação Coletiva com Planos Distintos nos Itens
                Call mfImpCabec_ROUBOCOT2(pNumArqDados, pobjstcPed, lobjstcPedLoc, pNumDe + Int(lobjstcPedLoc.Num_Item) - 1, pNumAte)
                Print #pNumArqDados, " "
                lCountLine = lCountLine + 1
                Print #pNumArqDados, "### CLÁUSULAS E CONDIÇÕES ###"
                lCountLine = lCountLine + 1
            End If
            lAuxiliar2 = lobjstcClau.gDescClausula
            Print #pNumArqDados, lAuxiliar2
            lCountLine = lCountLine + 1
            lTemInfo = True
        Next lobjstcClau

        Do While lCountLine < 85
            Print #pNumArqDados, " "
            lCountLine = lCountLine + 1
        Loop
    Next lobjstcPedLoc

    Call gM46V999.gpFecha2(lbdTab_Tipo_Fator)

    mfCriDados_ROUBOCOT5 = True

End Function

Public Function mfVerDigBens()
    Dim lI As Integer    'Utilizado no for...loop.
'Verificar se foi digitado Bens

    mfVerDigBens = False

    For lI = 1 To frmT46V129A.flxBens.Rows - 1
        If gM46V999.gfPreenchido(frmT46V129A.flxBens.TextMatrix(lI, 2)) And Trim(frmT46V129A.flxBens.TextMatrix(lI, 1)) <> " " Then
            mfVerDigBens = True
            Exit For
        End If
    Next lI
    Exit Function

ControleErro:
    mfVerDigBens = False

End Function

Private Function mfCriDados_MULTCOT1_Fonte12(ByVal pNumArqDados As Integer, ByVal pobjstcPed As stcA46V702B, _
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
    Dim lCountLine As Integer    'Contador de linhas.
    Dim lobjstcCob As stcA46V704B    'Objeto da classe (stc) de cobertura.
    Dim lobjstcSeg As stcA46V716B    'Outra seguradora.
    Dim lobjstcJur As stcA46V711B    'Objeto da classe (stc) de juros.
    Dim lMensagem As String
    Dim lValorParc As Variant
    Dim lDescFranquias() As String
    Dim lDescImportante() As String
    Dim lPrazo As Integer
    Dim lCodCobert As Integer
    Dim lNumMeses As Integer
    Dim lDescExp As Double
    Dim lbdTbm_Regiao As ADODB.Connection
    Dim lrsTbm_Regiao As clsYasRecordSet
    Dim lNumArqCXY As Integer
    Dim lPosX As Double
    Dim lPosY As Double
    Dim lLinCXY As Long


    'ROTINA CRIA CXY
    lCountPage = 1
    If gfIsMetaframe() Then
        mEndConfig = "L:\SYAS.CXY"

        Call gM46V999.gpGraLog(0, "Eh Metaframe forcou o L " & mEndConfig)
    Else
        mEndConfig = gM46V999.gAppPath & "SYAS.CXY"
    End If

    PodeGerarCYX = True
    lNumArqCXY = FreeFile

    On Error Resume Next
    Open mEndConfig For Output Access Write Lock Read Write As #lNumArqCXY

    lPosX = 0.6
    lPosY = 5

    Print #lNumArqCXY, "R Orientacao=Retrato; Papel=A4 210 x 297 mm; Altura=29.70; Largura=21.00; Vazio=0.00; ImgY=0.50; ImgX=0.50; ImgAltura=29.70; ImgLargura=21.00"

    lQtdPaginas = 1
    GeraCYX_PROP2 lLinCXY, "Página", 1, 18.5, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Título", 1, 10.3, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Processo Susep", 1.4, 10.7, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Emissão", 1.9, 2.1, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Validade", 1.9, 9.8, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Siscota", 2.4, 14.8, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Corretor", 3.1, 2, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Numero do pedido", 3.5, 2.4, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Número do item", 3.5, 6.4, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Tarifa", 3.5, 9, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Nosso numero", 3.5, 14.8, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Proponente - nome", 5.1, 0.6, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Proponente - codigo", 5.1, 12.1, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Proponente - tipo de pessoa", 5.1, 16.5, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Proponente - CNPJ/CPF - título", 5.9, 0.6, False, "N", lNumArqCXY
    '    GeraCYX_PROP2 lLinCXY, "Proponente - cat.registro - título", 5.9, 12.1, False, "N", lNumArqCXY
    '    GeraCYX_PROP2 lLinCXY, "Proponente - número registro - título", 5.9, 16.5, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Proponente - CNPJ/CPF e identificação", 6.5, 0.6, False, "N", lNumArqCXY
    '    GeraCYX_PROP2 lLinCXY, "Proponente - sigla resgistro", 6.5, 12.1, False, "N", lNumArqCXY
    '    GeraCYX_PROP2 lLinCXY, "Proponente - número registro", 6.5, 16.5, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Local em risco - logradouro", 8, 0.6, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Local em risco - numero", 8, 9, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Local em risco - complemento", 8, 12.5, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Local em risco - bairro", 8.8, 0.6, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Local em risco - CEP", 8.8, 5.8, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Local em risco - cidade", 8.8, 8.6, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Local em risco - UF", 8.8, 13.6, False, "N", lNumArqCXY
    'Ficha 1000305
    GeraCYX_PROP2 lLinCXY, "Dados do seguro - construção", 8.8, 14.4, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Dados do seguro - bens compreendidos", 8.8, 17.1, False, "N", lNumArqCXY

    GeraCYX_PROP2 lLinCXY, "Dados do seguro - classificação", 9.7, 0.6, False, "N", lNumArqCXY
    'Ficha 1000305
    GeraCYX_PROP2 lLinCXY, "Dados do seguro - valor em risco", 9.7, 14.4, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Dados do seguro - plano", 10.1, 0.6, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Dados do seguro - região", 10.8, 0.6, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Dados do seguro - Ass24h", 10.8, 13.7, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Dados do seguro - vigência", 11.6, 0.6, False, "N", lNumArqCXY

    '1000347 - Inclusão de informações da congênere em renovações
    If pobjstcPedLoc.Tip_Emissao = e_TipEmissao_RenCongenere Then
        GeraCYX_PROP2 lLinCXY, "Dados do seguro - tipo de emissão", 11.2, 13.3, False, "N", lNumArqCXY
        GeraCYX_PROP2 lLinCXY, "Congenere", 11.7, 10.7, False, "N", lNumArqCXY
        GeraCYX_PROP2 lLinCXY, "Apolice/Item", 12.2, 10.7, False, "N", lNumArqCXY
    Else
        GeraCYX_PROP2 lLinCXY, "Dados do seguro - tipo de emissão", 11.6, 10.7, False, "N", lNumArqCXY
    End If

    GeraCYX_PROP2 lLinCXY, "Dados para inspeção - nome da pessoa", 13.3, 5.3, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Dados para inspeção - telefone", 13.7, 4, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Detalhes usuário", 28.5, 0.6, False, "N", lNumArqCXY


    mfCriDados_MULTCOT1_Fonte12 = False

    Print #pNumArqDados, "Página: " & lQtdPaginas

    '001. Título do seguro
    lAuxiliar2 = pobjstcPed.getRpt.TituloSeguro

    Do While Len(lAuxiliar2) < 36
        lAuxiliar2 = " " & lAuxiliar2 & " "
    Loop
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '002. Processo Susep
    lAuxiliar2 = pobjstcPed.getRpt.ProcessoSusep
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '003. Emissão.
    Print #pNumArqDados, pobjstcPed.getRpt.Emissao

    '004. Validade.
    Print #pNumArqDados, pobjstcPed.getRpt.Validade

    '005. Siscota
    Print #pNumArqDados, pobjstcPed.getRpt.Siscota

    '006. Corretor (código e nome) .
    lAuxiliar1 = pobjstcPed.getRpt.CorretorCod
    lAuxiliar2 = pobjstcPed.getRpt.CorretorNome

    Print #pNumArqDados, Format$(lAuxiliar1, "0000") & " - " & lAuxiliar2

    '007. Número do pedido .
    Print #pNumArqDados, pobjstcPed.getRpt.Referencia

    '008. Número do item .
    Print #pNumArqDados, pobjstcPedLoc.getRpt.NumItem

    '009. Tarifa .
    Print #pNumArqDados, pobjstcPed.getRpt.Tarifa

    '010. Nosso número .
    Print #pNumArqDados, pobjstcPed.getRpt.NossoNumero

    '011. Proponente - nome do segurado .
    Print #pNumArqDados, pobjstcPed.getRpt.NomSegurado

    '012. Proponente - código do segurado - se houver .
    Print #pNumArqDados, pobjstcPed.getRpt.CodSegurado

    '013. Proponente - tipo de pessoa .
    Print #pNumArqDados, pobjstcPed.getRpt.TipoPessoa

    '014. Proponente - CNPJ ou CPF .
    lAuxiliar2 = pobjstcPed.getRpt.CpfCnpj
    lAuxiliar2 = lAuxiliar2 & pobjstcPed.getRpt.Identificacao1
    Print #pNumArqDados, lAuxiliar2

    '015. Identificação
    Print #pNumArqDados, pobjstcPed.getRpt.Identificacao2

    '016. Local em risco - logradouro .
    Print #pNumArqDados, pobjstcPedLoc.getRpt.LogrRisco

    '017. Local em risco - número .
    Print #pNumArqDados, pobjstcPedLoc.getRpt.NumLogrRisco

    '018. Local em risco - complemento .
    Print #pNumArqDados, pobjstcPedLoc.getRpt.DscComplRisco

    '019. Local em risco - bairro .
    Print #pNumArqDados, pobjstcPedLoc.getRpt.NomBairroRisco

    '020. Local em risco - CEP .
    Print #pNumArqDados, pobjstcPedLoc.getRpt.CepRisco

    '021. Local em risco - cidade .
    Print #pNumArqDados, pobjstcPedLoc.getRpt.CidadeRisco

    '022. Local em risco - UF .
    Print #pNumArqDados, pobjstcPedLoc.getRpt.UfRisco

    '023. Dados do seguro - construção .
    Print #pNumArqDados, pobjstcPedLoc.getRpt.TipoConstrucao

    '023B. Dados do seguro - Bens Compreendidos
    'Ficha 1000305
    Print #pNumArqDados, pobjstcPedLoc.getRpt.BensCompreendidos

    '024. Dados do seguro - classificação .
    Print #pNumArqDados, pobjstcPedLoc.getRpt.Classificacao

    '025. Dados do seguro - valor em risco .
    Print #pNumArqDados, pobjstcPedLoc.getRpt.ValorRisco

    'Dados do seguro - Plano / Produto
    Print #pNumArqDados, pobjstcPedLoc.getRpt.PlanoProduto

    '026. Dados do seguro - região .
    Print #pNumArqDados, pobjstcPedLoc.getRpt.Regiao

    Print #pNumArqDados, pobjstcPedLoc.getRpt.descricao_Ass24h

    '027. Dados do seguro - vigência .
    Print #pNumArqDados, pobjstcPed.getRpt.Vigencia

    '028. Dados do seguro - tipo de emissão .
    Print #pNumArqDados, pobjstcPedLoc.getRpt.TipoEmissao
    If pobjstcPedLoc.Tip_Emissao = e_TipEmissao_RenCongenere Then
        '1000347 - Inclusão de informações da congênere em renovações
        Print #pNumArqDados, pobjstcPedLoc.getRpt.dadosCongenere1
        Print #pNumArqDados, pobjstcPedLoc.getRpt.dadosCongenere2
    End If

    '029. Dados para inspeção - nome da pessoa para contato .
    Print #pNumArqDados, pobjstcPedLoc.getRpt.InspPessoaContato

    '030. Dados para inspeção - telefone para contato .
    Print #pNumArqDados, pobjstcPedLoc.getRpt.InspNumTel

    '031. Usuário
    Print #pNumArqDados, pobjstcPed.getRpt.UsuarioLiberacao


    lPosX = 0.6
    lPosY = 14.5
    lCountPage = 1
    lListaArq = mEndArq & "|" & mEndConfig & "|" & gM46V999.gEndImagCot1 & ";"

    'COBERTURAS
    lLinCXY = 15
    Call mpCriDados_CoberturasContratadas(pNumArqDados, lLinCXY, lPosY, lPosX, lNumArqCXY, pobjstcPed, pobjstcPedLoc)

    'FRANQUIAS - informar a cobertura e a respectiva franquia. Caso as coberturas contratadas não tenha franquia, informar texto "Não aplicável"
    Call mpCriDados_Franquia(pNumArqDados, lLinCXY, lPosY, lPosX, lNumArqCXY, pobjstcPed, pobjstcPedLoc)

    If pobjstcPed.ValTotalComissao <> 0 Then
        'DEMONSTRATIVO DO PRÊMIO
        Call mpCriDados_DemonstrativoPremios(pNumArqDados, lLinCXY, lPosY, lPosX, lNumArqCXY, pobjstcPed, pobjstcPedLoc)

        'Demonstrativo do parcelamento
        Call mpCriDados_DemonstrativoParcelamento(pNumArqDados, lLinCXY, lPosY, lPosX, lNumArqCXY, pobjstcPed, pobjstcPedLoc)
    End If
    'PERÍODO INDENITÁRIO - somente imprimir caso seja contratado (informar a cobertura e o período contratado).
    Call mpCriDados_PeriodoIndenitario(pNumArqDados, lLinCXY, lPosY, lPosX, lNumArqCXY, pobjstcPed, pobjstcPedLoc)

    'QUESTIONÁRIO - imprimir os itens sempre, mesmo que não tenha sido selecionado nenhum item.
    Call mpCriDados_Questionario(pNumArqDados, lLinCXY, lPosY, lPosX, lNumArqCXY, pobjstcPed, pobjstcPedLoc)
    'RELAÇÃO DE BENS - informar os bens informados. Caso não tenha, informar texto "Não informado"

    Call mpCriDados_RelacaoBens(pNumArqDados, lLinCXY, lPosY, lPosX, lNumArqCXY, pobjstcPed, pobjstcPedLoc)

    '1100589 - Tarifa 1014 e 1015 Julho 2011
    If gM46V999.gCodProduto >= 1014 Then
        Call mpImpClausulas(pNumArqDados, lLinCXY, lPosY, lPosX, lNumArqCXY, pobjstcPed, pobjstcPedLoc)
    End If

    'OBSERVAÇÕES - imprimir os textos fixos e texto digitado pelo analista no Syas.
    Call mpCriDados_Observacao(pNumArqDados, lLinCXY, lPosY, lPosX, lNumArqCXY, pobjstcPed, pobjstcPedLoc)

    '1100023 - Imprimir Informações de Cosseguro na Cotação
    Call mpCriDados_Cosseguro(pNumArqDados, lLinCXY, lPosY, lPosX, lNumArqCXY, pobjstcPed, pobjstcPedLoc)

    Call mpCriDados_UsoInterno(pNumArqDados, lLinCXY, lPosY, lPosX, lNumArqCXY, pobjstcPed, pobjstcPedLoc)

    Close #lNumArqCXY
    mfCriDados_MULTCOT1_Fonte12 = True
End Function


'1100589 - Tarifa 1014 e 1015 Julho 2011
Private Sub mpImpClausulas(pNumArqDados As Integer, pLinhaCXY As Long, pPosY As Double, pPosX As Double, pNumArqCXY As Integer, pobjstcPed As stcA46V702B, pobjstcPedLoc As stcA46V708B, Optional Titulo As String = "", Optional isLMI As Boolean = False)
    Dim lTextoClau As String
    Dim lobjstcClau As stcA46V703B
    Dim lPosicao As Integer
    Dim lTextoQuebra As String

    If Not isLMI Then
        If pLinhaCXY > 40 Then
            Call mpNovaPag(pLinhaCXY, pPosY, pPosX, pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc)
        End If
    End If

    If pobjstcPedLoc.ColecaoPedClau.Count > 0 Then
        'Titulo
        If Not isLMI Then
            GeraCYX_PROP2 pLinhaCXY, "Título Clausulas", pPosY, pPosX, True, "S", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
        Else
            GeraCYX_LMI pLinhaCXY, "Título Clausulas", pPosY, pPosX, True, "S", pNumArqCXY, pNumArqDados, pobjstcPed
        End If

        If Titulo = "" Then
            Print #pNumArqDados, String(80, "-") & " CLÁUSULAS " & String(80, "-")
        Else
            Print #pNumArqDados, Titulo
        End If

        'Clausulas
        For Each lobjstcClau In pobjstcPedLoc.ColecaoPedClau
            If Not isLMI Then
                GeraCYX_PROP2 pLinhaCXY, "Clausula " & lobjstcClau.getRpt.CodClausula, pPosY, pPosX, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
            Else
                GeraCYX_LMI pLinhaCXY, "Clausula " & lobjstcClau.getRpt.CodClausula, pPosY, pPosX, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed
            End If
            Print #pNumArqDados, lobjstcClau.getRpt.CodClausula & " - " & lobjstcClau.getRpt.DescClausula
            lTextoClau = ""
            lTextoClau = lobjstcClau.getRpt.TextoClausula
            'Texto da Clausula
            If Len(lTextoClau) > 0 Then
                Do
                    'encontra primeiro espaço antes do tamanho determinado
                    lPosicao = InStrRev(lTextoClau, " ", 110, vbTextCompare)
                    If lPosicao = 0 Then
                        If Not isLMI Then
                            GeraCYX_PROP2 pLinhaCXY, "Descrição Clausula " & lobjstcClau.getRpt.CodClausula, pPosY, pPosX, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
                        Else
                            GeraCYX_LMI pLinhaCXY, "Descrição Clausula " & lobjstcClau.getRpt.CodClausula, pPosY, pPosX, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed
                        End If
                        Print #pNumArqDados, "      " & lTextoClau
                        Exit Do
                    End If
                    '                    Debug.Print lPosicao
                    'guarda texto até o primeiro espaço antes do tamanho da linha definido
                    lTextoQuebra = Mid(lTextoClau, 1, lPosicao)
                    '                    Debug.Print lTextoQuebra
                    'retira da variável o texto que já foi separado acima
                    lTextoClau = Mid(lTextoClau, lPosicao, Len(lTextoClau))
                    '                    Debug.Print lTextoClau
                    If Not isLMI Then
                        GeraCYX_PROP2 pLinhaCXY, "Descrição Clausula " & lobjstcClau.getRpt.CodClausula, pPosY, pPosX, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
                    Else
                        GeraCYX_LMI pLinhaCXY, "Descrição Clausula " & lobjstcClau.getRpt.CodClausula, pPosY, pPosX, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed
                    End If
                    Print #pNumArqDados, "      " & lTextoQuebra
                Loop Until lPosicao = 0
            End If
        Next lobjstcClau
    End If

End Sub


Private Function mfCriDados_MULTPRO1_Fonte12(ByVal pNumArqDados As Integer, ByVal pobjstcPed As stcA46V702B, _
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
    Dim lPrazo As Integer
    Dim lCodCobert As Integer
    Dim lNumMeses As Integer
    Dim lDescCom As Double
    Dim lDescExp As Double
    Dim lValorParc As Variant
    Dim lMensagem As String
    Dim lOrdem As String    'Ordem.
    Dim lPriParcela As String   'Primeira parcela.
    Dim lDemParcelas As String    'Demais parcelas.
    Dim lJuros As String    'Juros.
    Dim lbdTbm_Regiao As ADODB.Connection
    Dim lrsTbm_Regiao As clsYasRecordSet
    Dim lNumArqCXY As Integer
    Dim lPosX As Double
    Dim lPosY As Double
    Dim lLinCXY As Long

    mfCriDados_MULTPRO1_Fonte12 = False

    'ROTINA CRIA CXY
    lCountPage = 1
    If gfIsMetaframe() Then
        mEndConfig = "L:\SYAS.CXY"

        Call gM46V999.gpGraLog(0, "Eh Metaframe forcou o L " & mEndConfig)
    Else
        mEndConfig = gM46V999.gAppPath & "SYAS.CXY"
    End If

    PodeGerarCYX = True
    lNumArqCXY = FreeFile

    '    On Error Resume Next
    Open mEndConfig For Output Access Write Lock Read Write As #lNumArqCXY

    lPosX = 0.6
    lPosY = 5

    Print #lNumArqCXY, "R Orientacao=Retrato; Papel=A4 210 x 297 mm; Altura=29.70; Largura=21.00; Vazio=0.00; ImgY=0.50; ImgX=0.50; ImgAltura=29.70; ImgLargura=21.00"

    lQtdPaginas = 1
    GeraCYX_PROP2 lLinCXY, "Página", 1, 18.5, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Título", 1, 10.3, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Processo Susep", 1.4, 10.3, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Emissão", 1.9, 2.1, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Siscota", 2.4, 14.8, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Corretor", 3.1, 2.1, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Número do pedido", 3.5, 2.5, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Número do item", 3.5, 6.4, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Tarifa", 3.5, 9, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Nosso Numero", 3.5, 14.8, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "DV", 4.1, 14.8, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Proponente - Nome", 5.1, 0.6, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Proponente - Código", 5.1, 12.1, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Proponente - tipo de pessoa", 5.1, 16.5, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Proponente - Identificação ", 5.9, 0.6, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Proponente - identificação 2", 6.3, 0.6, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Proponente - endereço", 7.2, 0.6, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Proponente - CEP", 6.75, 13, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Proponente - telefone", 6.75, 17.5, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Proponente - bairro", 8, 0.6, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Proponente - cidade", 8, 8.1, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Proponente - UF", 8, 14.4, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Local em risco - logradouro", 9.6, 0.6, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Local em risco - número", 9.6, 9.8, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Local em risco - complemento", 9.6, 12.5, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Local em risco - bairro", 10.4, 0.6, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Local em risco - CEP", 10.4, 5.8, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Local em risco - cidade", 10.4, 8.5, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Local em risco - UF", 10.4, 13.6, False, "N", lNumArqCXY
    'Ficha 1000305
    GeraCYX_PROP2 lLinCXY, "Dados do seguro - construção", 10.4, 14.4, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Dados do seguro - bens compreendidos", 10.4, 17.1, False, "N", lNumArqCXY

    GeraCYX_PROP2 lLinCXY, "Dados do seguro - classificação", 11.3, 0.6, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Dados do seguro - valor em risco", 11.3, 15.5, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Dados do seguro - plano", 11.7, 0.6, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Dados do seguro - regiao", 12.1, 1.9, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Dados do seguro - Ass24h", 12.1, 13.8, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Dados do seguro - vigêcia", 13.3, 0.6, False, "N", lNumArqCXY

    '1000347 - Inclusão de informações da congênere em renovações
    If pobjstcPedLoc.Tip_Emissao = e_TipEmissao_RenCongenere Then
        GeraCYX_PROP2 lLinCXY, "Dados do seguro - tipo de emissão", 12.8, 13.3, False, "N", lNumArqCXY
        GeraCYX_PROP2 lLinCXY, "Congenere", 13.3, 10.7, False, "N", lNumArqCXY
        GeraCYX_PROP2 lLinCXY, "Apolice/Item", 13.8, 10.7, False, "N", lNumArqCXY
    Else
        GeraCYX_PROP2 lLinCXY, "Dados do seguro - tipo de emissão", 13.3, 10.8, False, "N", lNumArqCXY
    End If

    GeraCYX_PROP2 lLinCXY, "Dados para inspeção - nome da pessoa", 14.9, 5.3, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Dados para inspeção - telefone", 15.3, 4.5, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Detalhe usuário", 28.08, 0.6, False, "N", lNumArqCXY

    Print #pNumArqDados, "Página: " & lQtdPaginas

    '001. Título do seguro (de acordo com a base de dados).
    '     Neste caso deve ser colocado PROPOSTA DE SEGURO YASUDA...
    '     RESIDENCIAL, CONDOMÍNIO ou EMPRESARIAL.
    '**********
    'Print #pNumArqDados, "PROPOSTA DE SEGURO YASUDA " & UCase(pObjstcPed.gDesc_Ramo)
    '**********
    Print #pNumArqDados, pobjstcPed.getRpt.DescRamo

    Print #pNumArqDados, pobjstcPed.getRpt.ProcessoSusep

    '003. Emissão.
    Print #pNumArqDados, pobjstcPed.getRpt.Emissao

    '005. Siscota
    Print #pNumArqDados, pobjstcPed.getRpt.Siscota

    '006. Corretor (código e nome) .
    lAuxiliar1 = pobjstcPed.getRpt.CorretorCod
    lAuxiliar2 = pobjstcPed.getRpt.CorretorNome

    Print #pNumArqDados, Format$(lAuxiliar1, "0000") & " - " & lAuxiliar2

    '007. Número do pedido .
    Print #pNumArqDados, pobjstcPed.getRpt.Referencia

    '008. Número do item .
    Print #pNumArqDados, pobjstcPedLoc.getRpt.NumItem

    '009. Tarifa .
    Print #pNumArqDados, pobjstcPed.getRpt.Tarifa

    '010. Nosso número .
    Print #pNumArqDados, pobjstcPed.getRpt.NossoNumero

    '011. Dígito Verificador (Código Unicidade)
    Print #pNumArqDados, pobjstcPed.getRpt.DV

    '012. Proponente - nome do segurado .
    Print #pNumArqDados, pobjstcPed.getRpt.NomSegurado

    '013. Proponente - código do segurado - se houver .
    Print #pNumArqDados, pobjstcPed.getRpt.CodSegurado

    '014. Proponente - tipo de pessoa .
    Print #pNumArqDados, pobjstcPed.getRpt.TipoPessoa

    '014. Proponente - CNPJ ou CPF .
    lAuxiliar2 = pobjstcPed.getRpt.CpfCnpj & "   " & pobjstcPed.getRpt.Identificacao1
    Print #pNumArqDados, lAuxiliar2

    '015. Identificação
    Print #pNumArqDados, pobjstcPed.getRpt.Identificacao2

    '017. Proponente - endereço .
    lAuxiliar2 = pobjstcPed.getRpt.Ped_Endereco
    'Ficha 900417
    Do While Len(lAuxiliar2) < 70
        lAuxiliar2 = lAuxiliar2 & " "
    Loop

    'CR7790 - A LABEL PROFISSAO SÓ DEVERA APARECER QUANDO FOR PESSOA FISICA
    'Profissao
    If UCase(Left(pobjstcPed.getRpt.TipoPessoa, 1)) = "F" Then
        lAuxiliar2 = lAuxiliar2 & pobjstcPed.getRpt.Profissao
        Print #pNumArqDados, lAuxiliar2
    Else
        lAuxiliar2 = lAuxiliar2 & " "
        Print #pNumArqDados, lAuxiliar2
    End If

    '018. Proponente - CEP .
    Print #pNumArqDados, pobjstcPed.getRpt.Ped_CEP

    '019. Proponente - Telefone .
    Print #pNumArqDados, pobjstcPed.getRpt.Ped_Fone

    '020. Proponente - bairro .
    Print #pNumArqDados, pobjstcPed.getRpt.Ped_Bairro

    '021. Proponente - cidade .
    Print #pNumArqDados, pobjstcPed.getRpt.Ped_Cidade

    '022. Proponente - UF .
    Print #pNumArqDados, pobjstcPed.getRpt.Ped_UF

    '023. Local em risco - logradouro .
    Print #pNumArqDados, pobjstcPedLoc.getRpt.LogrRisco

    '024. Local em risco - número .
    Print #pNumArqDados, pobjstcPedLoc.getRpt.NumLogrRisco

    '025. Local em risco - complemento .
    Print #pNumArqDados, pobjstcPedLoc.getRpt.DscComplRisco

    '026. Local em risco - bairro .
    Print #pNumArqDados, pobjstcPedLoc.getRpt.NomBairroRisco

    '027. Local em risco - CEP .
    Print #pNumArqDados, pobjstcPedLoc.getRpt.CepRisco

    '028. Local em risco - cidade .
    Print #pNumArqDados, pobjstcPedLoc.getRpt.CidadeRisco

    '029. Local em risco - UF .
    Print #pNumArqDados, pobjstcPedLoc.getRpt.UfRisco

    '030. Dados do seguro - construção .
    Print #pNumArqDados, pobjstcPedLoc.getRpt.TipoConstrucao

    'Ficha 1000305
    Print #pNumArqDados, pobjstcPedLoc.getRpt.BensCompreendidos

    '031. Dados do seguro - classificação .
    Print #pNumArqDados, pobjstcPedLoc.getRpt.Classificacao

    '032. Dados do seguro - valor em risco .
    Print #pNumArqDados, pobjstcPedLoc.getRpt.ValorRisco

    'Plano/Produto
    Print #pNumArqDados, pobjstcPedLoc.getRpt.PlanoProduto

    '033. Dados do seguro - região .
    Print #pNumArqDados, pobjstcPedLoc.getRpt.Regiao

    Print #pNumArqDados, pobjstcPedLoc.getRpt.descricao_Ass24h

    '034. Dados do seguro - vigência .
    Print #pNumArqDados, pobjstcPed.getRpt.Vigencia

    '028. Dados do seguro - tipo de emissão .
    Print #pNumArqDados, pobjstcPedLoc.getRpt.TipoEmissao
    If pobjstcPedLoc.Tip_Emissao = e_TipEmissao_RenCongenere Then
        '1000347 - Inclusão de informações da congênere em renovações
        Print #pNumArqDados, pobjstcPedLoc.getRpt.dadosCongenere1
        Print #pNumArqDados, pobjstcPedLoc.getRpt.dadosCongenere2
    End If

    '036. Dados para inspeção - nome da pessoa para contato .
    Print #pNumArqDados, pobjstcPedLoc.getRpt.InspPessoaContato

    '037. Dados para inspeção - telefone para contato .
    Print #pNumArqDados, pobjstcPedLoc.getRpt.InspNumTel

    '038. Usuário
    Print #pNumArqDados, pobjstcPed.getRpt.UsuarioLiberacao


    lPosX = 0.6
    lPosY = 16
    lCountPage = 1
    lListaArq = mEndArq & "|" & mEndConfig & "|" & gM46V999.gEndImagPro1 & ";"


    'COBERTURAS CONTRATADAS
    lLinCXY = 22
    Call mpCriDados_CoberturasContratadas(pNumArqDados, lLinCXY, lPosY, lPosX, lNumArqCXY, pobjstcPed, pobjstcPedLoc)

    'PERÍODO INDENITÁRIO - somente imprimir caso seja contratado (informar a cobertura e o período contratado).
    Call mpCriDados_PeriodoIndenitario(pNumArqDados, lLinCXY, lPosY, lPosX, lNumArqCXY, pobjstcPed, pobjstcPedLoc)

    'FRANQUIAS - informar a cobertura e a respectiva franquia. Caso as coberturas contratadas não tenha franquia, informar texto "Não aplicável"
    Call mpCriDados_Franquia(pNumArqDados, lLinCXY, lPosY, lPosX, lNumArqCXY, pobjstcPed, pobjstcPedLoc)

    'DEMONSTRATIVO DO PRÊMIO
    Call mpCriDados_DemonstrativoPremios(pNumArqDados, lLinCXY, lPosY, lPosX, lNumArqCXY, pobjstcPed, pobjstcPedLoc)

    'FORMA DE PAGAMENTO
    Call mpCriDados_FormaPagamento(pNumArqDados, lLinCXY, lPosY, lPosX, lNumArqCXY, pobjstcPed, pobjstcPedLoc)

    'QUESTIONÁRIO - imprimir os itens sempre, mesmo que não tenha sido selecionado nenhum item.
    Call mpCriDados_Questionario(pNumArqDados, lLinCXY, lPosY, lPosX, lNumArqCXY, pobjstcPed, pobjstcPedLoc)

    'RELAÇÃO DE BENS - informar os bens informados. Caso não tenha, informar texto "Não informado"
    Call mpCriDados_RelacaoBens(pNumArqDados, lLinCXY, lPosY, lPosX, lNumArqCXY, pobjstcPed, pobjstcPedLoc)

    'Relação de beneficiário
    Call mpCriDados_Beneficiario(pNumArqDados, lLinCXY, lPosY, lPosX, lNumArqCXY, pobjstcPed, pobjstcPedLoc)

    'DECLARAÇÕES
    Call mpCriDados_Declaracoes(pNumArqDados, lLinCXY, lPosY, lPosX, lNumArqCXY, pobjstcPed, pobjstcPedLoc)

    'LUCROS CESSANTES
    Call mpCriDados_LucrosCessantes(pNumArqDados, lLinCXY, lPosY, lPosX, lNumArqCXY, pobjstcPed, pobjstcPedLoc)

    '1100589 - Tarifa 1014 e 1015 Julho 2011
    If gM46V999.gCodProduto >= 1014 Then
        Call mpImpClausulas(pNumArqDados, lLinCXY, lPosY, lPosX, lNumArqCXY, pobjstcPed, pobjstcPedLoc)
    End If

    'OBSERVAÇÕES - imprimir os textos fixos e texto digitado pelo analista no Syas.
    Call mpCriDados_Observacao(pNumArqDados, lLinCXY, lPosY, lPosX, lNumArqCXY, pobjstcPed, pobjstcPedLoc)

    'TEXTO DA PROPOSTA
    Call mpCriDados_TextoProposta(pNumArqDados, lLinCXY, lPosY, lPosX, lNumArqCXY, pobjstcPed, pobjstcPedLoc)

    'DADOS PARA CORRESPONDÊNCIA
    Call mpCriDados_EnderecoCobranca(pNumArqDados, lLinCXY, lPosY, lPosX, lNumArqCXY, pobjstcPed, pobjstcPedLoc)

    'vencimento das parcelas
    Call mpCriDados_VencimentoParcelas(pNumArqDados, lLinCXY, lPosY, lPosX, lNumArqCXY, pobjstcPed, pobjstcPedLoc)

    'cosseguro
    Call mpCriDados_Cosseguro(pNumArqDados, lLinCXY, lPosY, lPosX, lNumArqCXY, pobjstcPed, pobjstcPedLoc)

    'USO DA SEGURADORA
    Call mpCriDados_UsoInterno(pNumArqDados, lLinCXY, lPosY, lPosX, lNumArqCXY, pobjstcPed, pobjstcPedLoc)

    Close #lNumArqCXY
    mfCriDados_MULTPRO1_Fonte12 = True

End Function
Private Sub mpCriDados_EnderecoCobranca(pNumArqDados As Integer, pLinhaCXY As Long, pPosY As Double, pPosX As Double, pNumArqCXY As Integer, pobjstcPed As stcA46V702B, pobjstcPedLoc As stcA46V708B)
    Dim lAuxiliar1 As Double
    Dim lAuxiliar2 As String
    Dim lAuxiliar3 As String
    Dim lTamanho As Integer
    Dim lCountLine As Integer

    If pLinhaCXY > 39 Then
        Call mpNovaPag(pLinhaCXY, pPosY, pPosX, pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc)
    End If
    '035. Dados para correspondência
    lAuxiliar2 = String(63, "-") & " DADOS PARA CORRESPONDÊNCIA " & String(63, "-")
    Print #pNumArqDados, lAuxiliar2
    GeraCYX_PROP2 pLinhaCXY, "Titulo dados de correspondencia", pPosY, pPosX, True, "S", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc

    lAuxiliar2 = "Endereço:"
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    GeraCYX_PROP2 pLinhaCXY, "Endereço de correspondêcia 1 ", pPosY, pPosX, False, "S", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc

    lAuxiliar2 = pobjstcPed.getRpt.End_Corresp_Logradouro
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    GeraCYX_PROP2 pLinhaCXY, "Endereço de correspondêcia ", pPosY, 2.5, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc

    lAuxiliar2 = "Bairro:" & Space(50)
    lAuxiliar2 = lAuxiliar2 & "Cidade:" & Space(50)
    lAuxiliar2 = lAuxiliar2 & "UF:" & Space(10)
    lAuxiliar2 = lAuxiliar2 & "CEP:"
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    GeraCYX_PROP2 pLinhaCXY, "Endereço de correspondêcia 1 ", pPosY, pPosX, False, "S", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc


    lAuxiliar2 = pobjstcPed.getRpt.End_Corresp_Bairro
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    GeraCYX_PROP2 pLinhaCXY, "Bairro ", pPosY, 2, False, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc

    lAuxiliar2 = pobjstcPed.getRpt.End_Corresp_Cidade & Space(14)
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    GeraCYX_PROP2 pLinhaCXY, "Cidade ", pPosY, 8.5, False, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc

    lAuxiliar2 = pobjstcPed.getRpt.End_Corresp_UF & Space(10)
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    GeraCYX_PROP2 pLinhaCXY, "UF ", pPosY, 14.3, False, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc

    lAuxiliar2 = pobjstcPed.getRpt.End_Corresp_CEP
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    GeraCYX_PROP2 pLinhaCXY, "CEP ", pPosY, 16.2, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
End Sub
Private Sub mpCriDados_LucrosCessantes(pNumArqDados As Integer, pLinhaCXY As Long, pPosY As Double, pPosX As Double, pNumArqCXY As Integer, pobjstcPed As stcA46V702B, pobjstcPedLoc As stcA46V708B)
    Dim lAuxiliar1 As Double
    Dim lAuxiliar2 As String
    Dim lAuxiliar3 As String
    Dim lTamanho As Integer
    Dim lCountLine As Integer
    Dim lbdTbm_Cob_LucroCess As ADODB.Connection
    Dim lrsTbm_Cob_LucroCess As clsYasRecordSet

    Dim lMensagem As String
    Dim lobjstcCob As stcA46V704B
    Dim lSelect As String

    lCountLine = 0

    If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTbm_Cob_LucroCess, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If

    For Each lobjstcCob In pobjstcPedLoc.ColecaoPedCobert
        If lobjstcCob.Num_Item = pobjstcPedLoc.Num_Item Then
            '1 - Procura pela Cobertura na Tabela de Coeficiente
            'Se encontrar é deve imprimir
            lSelect = " Select *  from Tbm_Cob_LucroCess "
            lSelect = lSelect & " Where Cod_Ramo = " & lobjstcCob.Cod_Ramo
            lSelect = lSelect & " and Cod_Cobert = " & lobjstcCob.Cod_Cobert
            lSelect = lSelect & " and Cod_Tarif = " & pobjstcPed.Cod_Produto
            lSelect = lSelect & " and Cod_Sequencia <= " & pobjstcPed.Cod_Sequencia

            If gM46V999.gfObtRegistro(lbdTbm_Cob_LucroCess, lSelect, lrsTbm_Cob_LucroCess, lMensagem) = False Then
                Call gM46V999.gpFecha2(lbdTbm_Cob_LucroCess)
                Call gM46V999.gpGraLog(2, lMensagem)
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            '2 - Se encontrar permite cadastro do Valor de Risco de Lucro Cessantes
            If Not lrsTbm_Cob_LucroCess.EOF Then
                If lCountLine = 0 Then
                    '030. Lucros Cessantes
                    lAuxiliar2 = String(75, "-") & " LUCROS CESSANTES " & String(75, "-")
                    Print #pNumArqDados, lAuxiliar2
                    GeraCYX_PROP2 pLinhaCXY, "Título de lucros cessantes", pPosY, pPosX, True, "S", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
                End If
                lCountLine = lCountLine + 1
                lAuxiliar3 = "Cobertura: " & lobjstcCob.gNomeCobertura & " - Valor em Risco: "
                lAuxiliar2 = Left(lAuxiliar3, 110) & String(100 - Len(lAuxiliar3), ".") & "R$ " & Format(lobjstcCob.Val_Risco_Cob, "standard")
                Print #pNumArqDados, lAuxiliar2
                GeraCYX_PROP2 pLinhaCXY, "Título de lucros cessantes", pPosY, pPosX, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
                lCountLine = lCountLine + 1
                If pLinhaCXY = 1 Then lCountLine = 0
            End If
        End If
    Next lobjstcCob

End Sub
Private Sub mpCriDados_Declaracoes(pNumArqDados As Integer, pLinhaCXY As Long, pPosY As Double, pPosX As Double, pNumArqCXY As Integer, pobjstcPed As stcA46V702B, pobjstcPedLoc As stcA46V708B)

'Declarações
    GeraCYX_PROP2 pLinhaCXY, "Título de declarações", pPosY, pPosX, True, "S", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
    Print #pNumArqDados, String(78, "-") & " DECLARAÇÕES " & String(78, "-")

    Dim lobjstcTextoImpressao As stcTextoImpressao
    For Each lobjstcTextoImpressao In pobjstcPedLoc.getRpt.Declaracoes
        GeraCYX_PROP2 pLinhaCXY, "Declarações - " & lobjstcTextoImpressao.Sequencia, pPosY, pPosX, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
        Print #pNumArqDados, lobjstcTextoImpressao.Descricao & " "
    Next

End Sub
Private Sub mpCriDados_Beneficiario(pNumArqDados As Integer, pLinhaCXY As Long, pPosY As Double, pPosX As Double, pNumArqCXY As Integer, pobjstcPed As stcA46V702B, pobjstcPedLoc As stcA46V708B)
    Dim lobjstcTextoImpressao As stcTextoImpressao

    If pLinhaCXY > 39 Then
        Call mpNovaPag(pLinhaCXY, pPosY, pPosX, pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc)
    End If
    '029. Beneficiarios
    GeraCYX_PROP2 pLinhaCXY, "relação de beneficiarios", pPosY, pPosX, True, "S", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
    Print #pNumArqDados, String(78, "-") & " BENEFICIÁRIOS " & String(78, "-")

    For Each lobjstcTextoImpressao In pobjstcPed.getRpt.Beneficiarios
        GeraCYX_PROP2 pLinhaCXY, "Beneficiário - " & lobjstcTextoImpressao.Sequencia, pPosY, pPosX, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
        Print #pNumArqDados, lobjstcTextoImpressao.Descricao & " "
    Next

    If pobjstcPed.getRpt.Beneficiarios.Count = 0 Then
        GeraCYX_PROP2 pLinhaCXY, "Beneficiários", pPosY, pPosX, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
        Print #pNumArqDados, "Não informado"
    End If

End Sub
Private Sub mpCriDados_UsoInterno(pNumArqDados As Integer, pLinhaCXY As Long, pPosY As Double, pPosX As Double, pNumArqCXY As Integer, pobjstcPed As stcA46V702B, pobjstcPedLoc As stcA46V708B)

    If pLinhaCXY > 40 Then
        Call mpNovaPag(pLinhaCXY, pPosY, pPosX, pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc)
    End If
    '036. Uso interno
    Print #pNumArqDados, String(183, "-")
    GeraCYX_PROP2 pLinhaCXY, "Linha", pPosY, pPosX, True, "S", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc

    '037. Unidade produtiva e Produtos
    Print #pNumArqDados, pobjstcPed.getRpt.UsoInterno_Unidade
    GeraCYX_PROP2 pLinhaCXY, "Linha", pPosY, pPosX, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc

    '039. Corretor .
    'Print #pNumArqDados, pobjstcPed.getRpt.UsoInterno_Corretor
    'GeraCYX_PROP2 pLinhaCXY, "Linha", pPosY, pPosX, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc

End Sub
Private Sub mpCriDados_FormaPagamento(pNumArqDados As Integer, pLinhaCXY As Long, pPosY As Double, pPosX As Double, pNumArqCXY As Integer, pobjstcPed As stcA46V702B, pobjstcPedLoc As stcA46V708B)
    Dim lAuxiliar1 As Double
    Dim lAuxiliar2 As String
    Dim lAuxiliar3 As String
    Dim lTamanho As Integer
    Dim lValorParc As Variant
    Dim lMensagem As String

    If pLinhaCXY > 37 Then
        Call mpNovaPag(pLinhaCXY, pPosY, pPosX, pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc)
    End If

    lAuxiliar2 = String(73, "-") & " FORMAS DE PAGAMENTO " & String(73, "-")
    Print #pNumArqDados, lAuxiliar2
    GeraCYX_PROP2 pLinhaCXY, "Título de forma de pagamento", pPosY, pPosX, True, "S", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc

    If pobjstcPed.objstcCtrlEmis.Num_Bloq <> 0 Then
        lAuxiliar2 = "Número de parcelas:" & Space(20) & "Parcelas antecipadas:" & Space(20) & "Número do bloqueto:"
    Else
        lAuxiliar2 = "Número de parcelas:" & Space(20) & "Parcelas antecipadas:"
    End If
    Print #pNumArqDados, lAuxiliar2
    GeraCYX_PROP2 pLinhaCXY, "Título de forma de pagamento", pPosY, pPosX, False, "S", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc

    '017. Número de parcelas .
    lAuxiliar1 = pobjstcPed.getRpt.Qtd_Parc
    Print #pNumArqDados, lAuxiliar1
    GeraCYX_PROP2 pLinhaCXY, "Parcelas", pPosY, 5, False, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc

    '018. Parcelas antecipadas .
    Print #pNumArqDados, pobjstcPed.getRpt.Cod_Iden_Pag_Antec

    If pobjstcPed.objstcCtrlEmis.Num_Bloq <> 0 Then
        '019. Número do bloqueto.
        GeraCYX_PROP2 pLinhaCXY, "Pagamento antecipado", pPosY, 10, False, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc

        lAuxiliar2 = pobjstcPed.getRpt.Num_Bloq & "-" & pobjstcPed.getRpt.DIG_BLOQUETO
        Print #pNumArqDados, lAuxiliar2
        GeraCYX_PROP2 pLinhaCXY, "Bloqueto", pPosY, 16, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
    Else
        GeraCYX_PROP2 pLinhaCXY, "Pagamento antecipado", pPosY, 10, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
    End If

    lAuxiliar2 = "Juros(a.m.):" & Space(33) & "Juros de mora(a.m.):"
    Print #pNumArqDados, lAuxiliar2
    GeraCYX_PROP2 pLinhaCXY, "Título de forma de pagamento", pPosY, pPosX, False, "S", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc

    '020. Juros ao mês .
    Print #pNumArqDados, pobjstcPed.getRpt.Per_Juros_Neg
    GeraCYX_PROP2 pLinhaCXY, "Juros", pPosY, 3, False, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc

    '021. Juros de mora .
    Print #pNumArqDados, pobjstcPed.getRpt.Per_Juros_Mora
    GeraCYX_PROP2 pLinhaCXY, "Juros", pPosY, 10, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc

    If pobjstcPed.Cod_Bco_Cobr <> 0 And pobjstcPed.Cod_Agen_Bas_Cor = 0 Then
        lAuxiliar2 = "Banco:"
        Print #pNumArqDados, lAuxiliar2
        GeraCYX_PROP2 pLinhaCXY, "Banco título", pPosY, pPosX, False, "S", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc

        lAuxiliar2 = pobjstcPed.getRpt.Banco
        Print #pNumArqDados, lAuxiliar2
        GeraCYX_PROP2 pLinhaCXY, "Banco carne", pPosY, 2, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
    End If

    If pobjstcPed.Cod_Bco_Cobr <> 0 And pobjstcPed.Cod_Agen_Bas_Cor <> 0 Then
        lAuxiliar2 = "Banco:" & Space(41) & "Agência:" & Space(40) & "Conta corrente:"
        Print #pNumArqDados, lAuxiliar2
        GeraCYX_PROP2 pLinhaCXY, "Dados de débito", pPosY, pPosX, False, "S", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc

        lAuxiliar2 = pobjstcPed.Cod_Bco_Cobr
        Print #pNumArqDados, lAuxiliar2
        GeraCYX_PROP2 pLinhaCXY, "Banco débito", pPosY, 2, False, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc

        lAuxiliar2 = pobjstcPed.getRpt.AgenciaDig
        Print #pNumArqDados, lAuxiliar2
        GeraCYX_PROP2 pLinhaCXY, "Agência débito", pPosY, 8, False, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc

        lAuxiliar2 = pobjstcPed.getRpt.ContaCorrente
        Print #pNumArqDados, lAuxiliar2
        GeraCYX_PROP2 pLinhaCXY, "C/C débito", pPosY, 14.5, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
    End If

    lAuxiliar2 = "Primeira parcela:" & Space(79) & "Demais parcelas:"
    Print #pNumArqDados, lAuxiliar2
    GeraCYX_PROP2 pLinhaCXY, "Título de forma de pagamento", pPosY, pPosX, False, "S", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc

    'Calcula o fracionamento
    With pobjstcPed
        Call gM46V998.gfCalFracionamento(pobjstcPed, pobjstcPedLoc.Tip_Emissao, _
                                         lValorParc, lMensagem, gM46V111.gProJurIOF, .Cod_Reg, .Val_IOF, .Val_Parc_Min)
    End With

    '022. Forma de pagamento: primeira parcela.
    Print #pNumArqDados, pobjstcPed.getRpt.Val_PriParcela
    GeraCYX_PROP2 pLinhaCXY, "Primeira parcela", pPosY, 3.5, False, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc

    '022. Forma de pagamento: demais parcelas.
    Print #pNumArqDados, pobjstcPed.getRpt.val_DemParcelas
    GeraCYX_PROP2 pLinhaCXY, "Demais parcela", pPosY, 14.7, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc

End Sub
Private Sub mpCriDados_TextoProposta(pNumArqDados As Integer, pLinhaCXY As Long, pPosY As Double, pPosX As Double, pNumArqCXY As Integer, pobjstcPed As stcA46V702B, pobjstcPedLoc As stcA46V708B)
    Dim lAuxiliar1 As Double
    Dim lAuxiliar2 As String
    Dim lAuxiliar3 As String

    GeraCYX_PROP2 pLinhaCXY, "TextoProposta", pPosY, pPosX, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
    Print #pNumArqDados, String(184, "-")

    Dim lobjstcTextoImpressao As stcTextoImpressao
    For Each lobjstcTextoImpressao In pobjstcPed.getRpt.TextoProposta
        GeraCYX_PROP2 pLinhaCXY, "Texto Proposta - " & lobjstcTextoImpressao.Sequencia, pPosY, pPosX, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
        Print #pNumArqDados, lobjstcTextoImpressao.Descricao
    Next

End Sub
Private Sub mpCriDados_Observacao(pNumArqDados As Integer, pLinhaCXY As Long, pPosY As Double, pPosX As Double, pNumArqCXY As Integer, pobjstcPed As stcA46V702B, pobjstcPedLoc As stcA46V708B)
    Dim objColec As Collection
    Dim wlobjObs As stcA46V714B

    GeraCYX_PROP2 pLinhaCXY, "Título de Obervações", pPosY, pPosX, True, "S", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
    Print #pNumArqDados, String(78, "-") & " OBSERVAÇÕES " & String(78, "-")

    Set objColec = pobjstcPedLoc.getRpt.Observacoes()
    For Each wlobjObs In objColec
        GeraCYX_PROP2 pLinhaCXY, "Texto Proposta - " & wlobjObs.Class_ID, pPosY, pPosX, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
        Print #pNumArqDados, IIf(Len(wlobjObs.Observacao) = 0, " ", wlobjObs.Observacao)
    Next

    Set objColec = Nothing
    Set wlobjObs = Nothing
End Sub




Private Sub mpCriDados_RelacaoBens(pNumArqDados As Integer, pLinhaCXY As Long, pPosY As Double, pPosX As Double, pNumArqCXY As Integer, pobjstcPed As stcA46V702B, pobjstcPedLoc As stcA46V708B)
    Dim lAuxiliar1 As Double
    Dim lAuxiliar2 As String
    Dim lAuxiliar3 As String
    Dim lTamanho As Integer
    Dim lCountLine As Integer
    Dim lobjstcBem As stcA46V715B    'Bens


    If pLinhaCXY > 40 Then
        Call mpNovaPag(pLinhaCXY, pPosY, pPosX, pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc)
    End If

    'Relação de Bens
    lCountLine = 0
    For Each lobjstcBem In pobjstcPedLoc.ColecaoPedLocBem
        If lCountLine = 0 Then
            lAuxiliar2 = String(75, "-") & " RELAÇÃO DE BENS " & String(75, "-")
            Print #pNumArqDados, lAuxiliar2
            GeraCYX_PROP2 pLinhaCXY, "Título de relação de bens", pPosY, pPosX, True, "S", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
        End If
        lCountLine = lCountLine + 1
        lAuxiliar3 = lobjstcBem.getRpt.TipBem & " - " & lobjstcBem.getRpt.DescBen
        lAuxiliar2 = Format(lCountLine, "000") & " - " & Left(lAuxiliar3, 100) & String(90 - Len(Left(lAuxiliar3, 90)), ".") & lobjstcBem.getRpt.Val_Bem
        Print #pNumArqDados, lAuxiliar2
        GeraCYX_PROP2 pLinhaCXY, "relação de bens " & lCountLine, pPosY, pPosX, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
        If pLinhaCXY = 1 Then lCountLine = 0
    Next lobjstcBem
    If lCountLine = 0 And pLinhaCXY <> 1 Then
        lAuxiliar2 = String(75, "-") & " RELAÇÃO DE BENS " & String(75, "-")
        Print #pNumArqDados, lAuxiliar2
        GeraCYX_PROP2 pLinhaCXY, "Título de relação de bens", pPosY, pPosX, True, "S", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc

        Print #pNumArqDados, "Não informado"
        GeraCYX_PROP2 pLinhaCXY, "relação de bens " & lCountLine, pPosY, pPosX, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
    End If
End Sub
Private Sub mpCriDados_Questionario(pNumArqDados As Integer, pLinhaCXY As Long, pPosY As Double, pPosX As Double, pNumArqCXY As Integer, pobjstcPed As stcA46V702B, pobjstcPedLoc As stcA46V708B)

    Dim lobjstcFator As stcA46V720B  'Objeto da classe de estrutura stcA46V720B (TAB_PED_FAT_RISCO).
    Dim objColec As Collection

    Set objColec = pobjstcPedLoc.getRpt.gFatoresDeProtecao
    If objColec.Count > 0 Then
        If pLinhaCXY > 35 Then
            Call mpNovaPag(pLinhaCXY, pPosY, pPosX, pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc)
        End If
        GeraCYX_PROP2 pLinhaCXY, "Título de questionário ", pPosY, pPosX, True, "S", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
        Print #pNumArqDados, String(78, "-") & " QUESTIONÁRIO " & String(78, "-")
        For Each lobjstcFator In objColec
            GeraCYX_PROP2 pLinhaCXY, "Questionário " & lobjstcFator.Cod_Fator, pPosY, pPosX, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
            Print #pNumArqDados, lobjstcFator.Desc_Fator
        Next
    End If
    Set objColec = Nothing

End Sub
Private Sub mpCriDados_Franquia(pNumArqDados As Integer, pLinhaCXY As Long, pPosY As Double, pPosX As Double, pNumArqCXY As Integer, pobjstcPed As stcA46V702B, pobjstcPedLoc As stcA46V708B)
    Dim lAuxiliar1 As Double
    Dim lAuxiliar2 As String
    Dim lAuxiliar3 As String
    Dim lTamanho As Integer
    Dim lCountLine As Integer
    Dim lobjstcCob As stcA46V704B
    Dim lDescFranquias() As String
    Dim lI As Integer
    Dim lTemInfo As Boolean

    '1200903 - Tarifa - 1046 e 1047 - Item 15, 16 e 17
    Dim blnFranqEspec_59_e_60 As Boolean
    blnFranqEspec_59_e_60 = False
    Dim blnFranqEspec_22_59_e_60 As Boolean
    blnFranqEspec_22_59_e_60 = False
    If pobjstcPedLoc.Cod_Iden_Tarif >= 1046 Then
        If Trim(pobjstcPedLoc.Cod_Clasf) = gM46V999.gfBuscarDePara("420", 2) Then  'Item 15 - Padarias
            blnFranqEspec_59_e_60 = True
        ElseIf Trim(pobjstcPedLoc.Cod_Clasf) = gM46V999.gfBuscarDePara("061", 2) Then  'Item 16 - Bares e Restaurantes
            blnFranqEspec_59_e_60 = True
        ElseIf (Trim(pobjstcPedLoc.Cod_Clasf) = gM46V999.gfBuscarDePara("282", 2) And _
                Trim(pobjstcPedLoc.Cod_Compl_Clasf) = gM46V999.gfBuscarDePara("282", 2, "10")) Then 'Item 16 - Hoteis e Pousadas
            blnFranqEspec_59_e_60 = True
        ElseIf (Trim(pobjstcPedLoc.Cod_Clasf) = gM46V999.gfBuscarDePara("282A", 2) And _
                Trim(pobjstcPedLoc.Cod_Compl_Clasf) = gM46V999.gfBuscarDePara("282A", 2, "10")) Then 'Item 16 - Hoteis e Pousadas
            blnFranqEspec_59_e_60 = True
        ElseIf (Trim(pobjstcPedLoc.Cod_Clasf) = gM46V999.gfBuscarDePara("496", 2) And _
                Trim(pobjstcPedLoc.Cod_Compl_Clasf) = gM46V999.gfBuscarDePara("496", 2, "10")) Then 'Item 17 - SuperMercados
            blnFranqEspec_22_59_e_60 = True
        Else
            blnFranqEspec_59_e_60 = False
            blnFranqEspec_22_59_e_60 = False
        End If
    End If

    If pLinhaCXY > 37 Then
        Call mpNovaPag(pLinhaCXY, pPosY, pPosX, pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc)
    End If
    lTemInfo = False
    '026. Franquias
    lCountLine = 0
    For Each lobjstcCob In pobjstcPedLoc.ColecaoPedCobert
        If (lobjstcCob.getRpt.temFranquia = True) Or _
           (blnFranqEspec_59_e_60 = True And (lobjstcCob.Cod_Cobert = gM46V999.gfBuscarDePara(59, 1) Or lobjstcCob.Cod_Cobert = gM46V999.gfBuscarDePara(60, 1))) Or _
           (blnFranqEspec_22_59_e_60 = True And (lobjstcCob.Cod_Cobert = gM46V999.gfBuscarDePara(22, 1) Or lobjstcCob.Cod_Cobert = gM46V999.gfBuscarDePara(59, 1) Or _
           lobjstcCob.Cod_Cobert = gM46V999.gfBuscarDePara(60, 1))) Then   '1200903 - Tarifa - 1046 e 1047 - Item 15, 16 e 17
            'cobertura
            lTemInfo = True
            If lCountLine = 0 Then
                lAuxiliar2 = String(81, "-") & " FRANQUIAS " & String(81, "-")
                GeraCYX_PROP2 pLinhaCXY, "Título franquia", pPosY, pPosX, True, "S", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
                Print #pNumArqDados, lAuxiliar2
            End If

            If pLinhaCXY = 42 Then
                Call mpNovaPag(pLinhaCXY, pPosY, pPosX, pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc)
            End If

            lCountLine = lCountLine + 1
            lAuxiliar2 = lobjstcCob.getRpt.codFranquia & "-" & lobjstcCob.getRpt.nomFranquia

            GeraCYX_PROP2 pLinhaCXY, "franquia " & lCountLine, pPosY, pPosX, False, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)


            lAuxiliar3 = lobjstcCob.getRpt.descFranquia
            If gM46V999.gfPreenchido(lAuxiliar3) Then
                GeraCYX_PROP2 pLinhaCXY, "franquia " & lCountLine, pPosY, 10, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
                Print #pNumArqDados, IIf(lAuxiliar3 = "", " ", lAuxiliar3)
            Else
                lAuxiliar3 = lobjstcCob.getRpt.descFranquiaPrazo
                If Trim(lAuxiliar3) <> "" Then
                    GeraCYX_PROP2 pLinhaCXY, "franquia " & lCountLine, pPosY, 10, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
                    Print #pNumArqDados, lAuxiliar3
                End If
            End If

            If pLinhaCXY = 1 Then lCountLine = 0
        End If
    Next lobjstcCob

    '1100352 - Tarifa 1010 e 1011 Maio 2011
    '    Call mpDescricaoFranquia_Corretor(pobjstcPedLoc, pobjstcPed, lDescFranquias, 50)

    '    For lI = 0 To 50
    '        If lDescFranquias(lI) = "" Then Exit For
    '        If lCountLine = 0 Then
    '            lAuxiliar2 = String(85, "-") & "FRANQUIAS" & String(85, "-")
    '            GeraCYX_PROP2 pLinhaCXY, "Título franquia", pPosY, pPosX, True, "S", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
    '            Print #pNumArqDados, lAuxiliar2
    '        End If
    '        lTemInfo = True
    '        lCountLine = lCountLine + 1
    '        lAuxiliar2 = lDescFranquias(lI)
    '        GeraCYX_PROP2 pLinhaCXY, "franquia " & lCountLine, pPosY, pPosX, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
    '        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    '        If pLinhaCXY = 1 Then lCountLine = 0
    '    Next lI

    'Call mpDescricaoFranquia(pobjstcPedLoc, pobjstcPed, lDescFranquias, 50)
    'For li = 0 To 50
    '   If lDescFranquias(li) = "" Then Exit For
    '   If lCountLine = 0 Then
    '       lAuxiliar2 = String(85, "-") & "FRANQUIAS" & String(85, "-")
    '       Print #pNumArqDados, lAuxiliar2
    '       GeraCYX_PROP2 pLinhaCXY, "Título franquia", pPosY, pPosX, True, "S", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
    '   End If
    '   lTemInfo = True
    '   lCountLine = lCountLine + 1
    '   lAuxiliar2 = lDescFranquias(li)
    '   Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    '   GeraCYX_PROP2 pLinhaCXY, "franquia " & lCountLine, pPosY, pPosX, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
    '   If pLinhaCXY = 1 Then lCountLine = 0
    'Next li

    If lTemInfo = False Then
        lAuxiliar2 = String(85, "-") & "FRANQUIAS" & String(85, "-")
        GeraCYX_PROP2 pLinhaCXY, "Título franquia", pPosY, pPosX, True, "S", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
        Print #pNumArqDados, lAuxiliar2
        GeraCYX_PROP2 pLinhaCXY, "franquia " & lCountLine, pPosY, pPosX, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
        Print #pNumArqDados, "Não aplicável"
    End If

End Sub
Private Sub mpCriDados_PeriodoIndenitario(pNumArqDados As Integer, pLinhaCXY As Long, pPosY As Double, pPosX As Double, pNumArqCXY As Integer, pobjstcPed As stcA46V702B, pobjstcPedLoc As stcA46V708B)
    Dim lAuxiliar1 As Double
    Dim lAuxiliar2 As String
    Dim lAuxiliar3 As String
    Dim lTamanho As Integer
    Dim lCountLine As Integer
    Dim lobjstcCob As stcA46V704B
    Dim lCodCobert As Integer
    Dim lPrazo As Integer

    '025. Período Indenitário
    lCountLine = 0
    For Each lobjstcCob In pobjstcPedLoc.ColecaoPedCobert
        lAuxiliar1 = IIf(IsNumeric(lobjstcCob.Num_Meses_PI), lobjstcCob.Num_Meses_PI, 0)
        If lAuxiliar1 > 0 Then
            If lCountLine = 0 Then
                If pLinhaCXY > 37 Then
                    Call mpNovaPag(pLinhaCXY, pPosY, pPosX, pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc)
                End If
                lAuxiliar2 = String(74, "-") & " PERÍODO INDENITÁRIO " & String(74, "-")
                Print #pNumArqDados, lAuxiliar2
                GeraCYX_PROP2 pLinhaCXY, "Título período indenitário", pPosY, pPosX, True, "S", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
            End If
            lCountLine = lCountLine + 1
            'cobertura
            Print #pNumArqDados, Format(lobjstcCob.getRpt.CodCobert, "000") & "-" & Left(lobjstcCob.getRpt.nomCobert, 60)
            GeraCYX_PROP2 pLinhaCXY, "período indenitário cob " & lCountLine, pPosY, pPosX, False, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
            Print #pNumArqDados, lobjstcCob.getRpt.Num_Meses_PI
            GeraCYX_PROP2 pLinhaCXY, "período indenitário PI " & lCountLine, pPosY, 12, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
            If pLinhaCXY = 1 Then lCountLine = 0
        End If
    Next lobjstcCob

End Sub
Private Sub mpCriDados_DemonstrativoParcelamento(pNumArqDados As Integer, pLinhaCXY As Long, pPosY As Double, pPosX As Double, pNumArqCXY As Integer, pobjstcPed As stcA46V702B, pobjstcPedLoc As stcA46V708B)
    Dim lAuxiliar1 As Double
    Dim lAuxiliar2 As String
    Dim lAuxiliar3 As String
    Dim lTamanho As Integer
    Dim lCountLine As Integer
    Dim lValorParc As Variant
    Dim lMensagem As String
    Dim lI As Integer
    Dim lOrdem As Integer
    Dim lPriParcela As String
    Dim lDemParcelas As String    'Demais parcelas.
    Dim lobjstcJur As stcA46V711B
    Dim lJuros As String
    Dim lParcelas(12)

    If pLinhaCXY > 29 Then
        Call mpNovaPag(pLinhaCXY, pPosY, pPosX, pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc)
    End If

    lAuxiliar2 = String(62, "-") & " DEMONSTRATIVO DE PARCELAMENTO " & String(62, "-")
    Print #pNumArqDados, lAuxiliar2
    GeraCYX_PROP2 pLinhaCXY, "Título de demonstrativo de parcelamento", pPosY, 0.6, True, "S", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc

    lAuxiliar2 = "                   Primeira parcela a vista                                                                      Primeira parcela a 30 dias"
    Print #pNumArqDados, lAuxiliar2
    GeraCYX_PROP2 pLinhaCXY, "Título de demonstrativo de parcelamento 2", pPosY, 0.6, True, "S", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc

    lAuxiliar2 = String(80, "-") & Space(22) & String(82, "-")
    Print #pNumArqDados, lAuxiliar2
    GeraCYX_PROP2 pLinhaCXY, "Título de demonstrativo de parcelamento 3", pPosY, 0.6, True, "S", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc

    lAuxiliar2 = "Parc              1ª parcela                 Demais           Juros                        1ª parcela            Demais           Juros "
    Print #pNumArqDados, lAuxiliar2
    GeraCYX_PROP2 pLinhaCXY, "Título de demonstrativo de parcelamento 4", pPosY, 0.6, True, "S", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc

    lCountLine = lCountLine + 4

    'Calcula o fracionamento
    With pobjstcPed
        Call gM46V998.gfCalFracionamento(pobjstcPed, pobjstcPedLoc.Tip_Emissao, _
                                         lValorParc, lMensagem, gM46V111.gProJurIOF, .Cod_Reg, .Val_IOF, .Val_Parc_Min)
    End With

    '018. Formas de pagamento - primeira parcela a 30 dias.
    'Ficha 900665
    lOrdem = 1
    If pobjstcPed.Qtd_Parc = 0 And pobjstcPed.VAL_PR <> 0 Then
        pobjstcPed.Qtd_Parc = 1
    End If
    For lI = 1 To 12
        If lI > pobjstcPed.Qtd_Parc Then
            '            Print #pNumArqDados, " "
            '            lCountLine = lCountLine + 1
            Exit For
        Else
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
                If lTamanho < 9 Then
                    lDemParcelas = Space(9 - lTamanho) & lDemParcelas
                End If
                'd) Juros.
                Set lobjstcJur = New stcA46V711B
                On Error Resume Next
                Set lobjstcJur = pobjstcPed.ColecaoPedJuros.Item(pobjstcPed.NOSSO_NUMERO & Format(lI, "00"))
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
                lAuxiliar2 = Format(lOrdem, "00") & "    " & lPriParcela & "  " & lDemParcelas & "  " & lJuros
            Else
                lAuxiliar2 = "- "
            End If
            If lValorParc(lI, 3) <> 0 Then
                If lAuxiliar2 = "- " Then   'Ficha 1000099 - Planos 2131 a 2138 Corretor Schroeder - para imprimir corretamente parc a 30dias
                    lAuxiliar2 = Format(lOrdem, "00") & Space(31)
                End If
                'b) Valor da primeira parcela.
                lAuxiliar1 = lValorParc(lI, 3)
                If lAuxiliar1 <> 0 Then
                    lPriParcela = Format$(lAuxiliar1, "standard")
                    lTamanho = Len(lPriParcela)
                    If lTamanho < 10 Then
                        lPriParcela = Space(10 - lTamanho) & lPriParcela
                    End If
                    'c) Valor das demais parcelas.
                    lAuxiliar1 = lValorParc(lI, 4)
                    lDemParcelas = Format$(lAuxiliar1, "standard")
                    lTamanho = Len(lDemParcelas)
                    If lTamanho < 9 Then
                        lDemParcelas = Space(9 - lTamanho) & lDemParcelas
                    End If
                    'd) Juros.
                    Set lobjstcJur = New stcA46V711B
                    On Error Resume Next
                    Set lobjstcJur = pobjstcPed.ColecaoPedJuros.Item(pobjstcPed.NOSSO_NUMERO & Format(lI, "00"))
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
                    lAuxiliar2 = lAuxiliar2 & "      " & lPriParcela & "  " & lDemParcelas & "  " & lJuros
                    lCountLine = lCountLine + 1
                Else
                    lAuxiliar2 = lAuxiliar2 & " "
                End If
            Else
                lAuxiliar2 = lAuxiliar2 & " "
            End If
            If lAuxiliar2 <> "-  " Then
                Print #pNumArqDados, lAuxiliar2
                GeraCYX_PROP2 pLinhaCXY, "Parcela ", pPosY, 0.6, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc, "COURIER NEW"
            End If
        End If
    Next lI
    If pobjstcPed.Qtd_Parc = 1 And pobjstcPed.VAL_PR <= 0 Then
        pobjstcPed.Qtd_Parc = 0
    End If
End Sub



Private Sub mpCriDados_DemonstrativoPremios(pNumArqDados As Integer, pLinhaCXY As Long, pPosY As Double, pPosX As Double, pNumArqCXY As Integer, pobjstcPed As stcA46V702B, pobjstcPedLoc As stcA46V708B)
    Dim lAuxiliar1 As Double
    Dim lAuxiliar2 As String
    Dim lAuxiliar3 As String
    Dim lTamanho As Integer

    If pLinhaCXY > 40 Then
        Call mpNovaPag(pLinhaCXY, pPosY, pPosX, pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc)
    End If
    lAuxiliar2 = String(70, "-") & "DEMONSTRATIVO DO PRÊMIO" & String(70, "-")
    Print #pNumArqDados, lAuxiliar2
    GeraCYX_PROP2 pLinhaCXY, "Título de demonstrativo de prêmio", pPosY, pPosX, True, "S", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc

    lAuxiliar2 = "Prêmio Líquido" & Space(20) & "Custo da Apólice" & Space(20) & "IOF (7,38%)" & Space(20) & "Prêmio Total Vista"
    Print #pNumArqDados, lAuxiliar2
    GeraCYX_PROP2 pLinhaCXY, "Título de demonstrativo de prêmio 2", pPosY, pPosX, True, "S", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc

    '013. Totalizar - Prêmio Líquido
    lAuxiliar3 = pobjstcPed.getRpt.ValPrLiqTotal

    lTamanho = Len(lAuxiliar3)
    If lTamanho < 14 Then
        lAuxiliar3 = Space(14 - lTamanho) & lAuxiliar3
    End If
    Print #pNumArqDados, lAuxiliar3
    GeraCYX_PROP2 pLinhaCXY, "Prêmio líquido", pPosY, pPosX, False, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc, "COURIER NEW"


    '014. Custo da Apólice
    lAuxiliar3 = pobjstcPed.getRpt.ValCusEmis
    lTamanho = Len(lAuxiliar3)
    If lTamanho < 14 Then
        lAuxiliar3 = Space(14 - lTamanho) & lAuxiliar3
    End If
    Print #pNumArqDados, lAuxiliar3
    GeraCYX_PROP2 pLinhaCXY, "Custo de apólice", pPosY, 5, False, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc, "COURIER NEW"

    '015. Valor IOF.
    lAuxiliar3 = pobjstcPed.getRpt.ValIOF
    lTamanho = Len(lAuxiliar3)
    If lTamanho < 14 Then
        lAuxiliar3 = Space(14 - lTamanho) & lAuxiliar3
    End If
    Print #pNumArqDados, lAuxiliar3
    GeraCYX_PROP2 pLinhaCXY, "IOF", pPosY, 9, False, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc, "COURIER NEW"

    '016. Prêmio total a vista.
    lAuxiliar3 = pobjstcPed.getRpt.ValPremioTotal
    lTamanho = Len(lAuxiliar3)
    If lTamanho < 14 Then
        lAuxiliar3 = Space(14 - lTamanho) & lAuxiliar3
    End If
    Print #pNumArqDados, lAuxiliar3
    GeraCYX_PROP2 pLinhaCXY, "Prêmio total", pPosY, 14, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc, "COURIER NEW"
End Sub

Private Sub mpCriDados_VencimentoParcelas(pNumArqDados As Integer, pLinhaCXY As Long, pPosY As Double, pPosX As Double, pNumArqCXY As Integer, pobjstcPed As stcA46V702B, pobjstcPedLoc As stcA46V708B)
    Dim lAuxiliar1 As Double
    Dim lAuxiliar2 As String
    Dim lAuxiliar3 As String
    Dim lCountLine As Integer
    Dim lObjstcVenc As stcA46V713B


    If pobjstcPed.ColecaoPedVenc.Count > 0 Then
        lAuxiliar2 = ""
        If pLinhaCXY > 40 Then
            Call mpNovaPag(pLinhaCXY, pPosY, pPosX, pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc)
        End If
        Print #pNumArqDados, String(50, "-") & "VENCIMENTO DAS PARCELAS" & String(50, "-")
        GeraCYX_PROP2 pLinhaCXY, "Título de vencimento", pPosY, pPosX, True, "S", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
        For Each lObjstcVenc In pobjstcPed.ColecaoPedVenc
            If lObjstcVenc.Num_Parc <= pobjstcPed.Qtd_Parc Then
                If lObjstcVenc.Num_Parc < 6 Then
                    lAuxiliar2 = lAuxiliar2 & "(" & lObjstcVenc.Num_Parc & "ª-" & gM46V999.gfForData(2, lObjstcVenc.DAT_VENC) & ") "
                Else
                    If lObjstcVenc.Num_Parc = 6 Then
                        Print #pNumArqDados, lAuxiliar2
                        GeraCYX_PROP2 pLinhaCXY, "vencimento", pPosY, pPosX, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
                        lAuxiliar2 = ""
                    End If
                    lAuxiliar2 = lAuxiliar2 & "(" & lObjstcVenc.Num_Parc & "ª-" & gM46V999.gfForData(2, lObjstcVenc.DAT_VENC) & ") "
                End If
            End If
        Next lObjstcVenc
        If pobjstcPed.Qtd_Parc < 6 Then
            Print #pNumArqDados, lAuxiliar2
            GeraCYX_PROP2 pLinhaCXY, "vencimento", pPosY, pPosX, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
        Else
            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
            GeraCYX_PROP2 pLinhaCXY, "vencimento", pPosY, pPosX, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
        End If
    End If

End Sub
Private Sub mpCriDados_Cosseguro(pNumArqDados As Integer, pLinhaCXY As Long, pPosY As Double, pPosX As Double, pNumArqCXY As Integer, pobjstcPed As stcA46V702B, pobjstcPedLoc As stcA46V708B)
    Dim lAuxiliar2 As String

    If pLinhaCXY > 40 Then
        Call mpNovaPag(pLinhaCXY, pPosY, pPosX, pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc)
    End If
    '032. Cosseguro
    If pobjstcPed.ColecaoPedCoss.Count > 0 Then
        lAuxiliar2 = String(80, "-") & " COSSEGURO " & String(80, "-")
        Print #pNumArqDados, lAuxiliar2
        GeraCYX_PROP2 pLinhaCXY, "Título Cosseguro", pPosY, pPosX, True, "S", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
        Print #pNumArqDados, pobjstcPed.getRpt.Cosseguro
        GeraCYX_PROP2 pLinhaCXY, "Cosseguro", pPosY, pPosX, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
    End If
End Sub
Private Sub mpCriDados_CoberturasContratadas(pNumArqDados As Integer, pLinhaCXY As Long, pPosY As Double, pPosX As Double, pNumArqCXY As Integer, pobjstcPed As stcA46V702B, pobjstcPedLoc As stcA46V708B)
    Dim lAuxiliar1 As Double
    Dim lAuxiliar2 As String
    Dim lAuxiliar3 As String
    Dim lCountLine As Integer
    Dim lobjstcCob As stcA46V704B
    Dim lTamanho As Integer
    Dim lObjstcCorr As stcA46V705B    'Objeto stcA46V705B  (corretor).


    '010. Coberturas contratadas
    lAuxiliar2 = String(66, "-") & " COBERTURAS CONTRATADAS " & String(66, "-")
    Print #pNumArqDados, lAuxiliar2
    GeraCYX_PROP2 pLinhaCXY, "Titulo coberturas", pPosY, pPosX, True, "S", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc

    lAuxiliar2 = "COBERTURA                                                                     LIMITE MÁXIMO DE GARANTIA(R$)" & IIf(pobjstcPed.ValTotalComissao = 0, "               PRÊMIO NET (R$)", "          PRÊMIO LÍQUIDO(R$)")  '1000894 - Imprimir Literal -Premio Net-
    Print #pNumArqDados, lAuxiliar2
    GeraCYX_PROP2 pLinhaCXY, "Titulo coberturas 2", pPosY, pPosX, True, "S", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
    lCountLine = 0
    For Each lobjstcCob In pobjstcPedLoc.ColecaoPedCobert
        'Código
        lCountLine = lCountLine + 1
        lAuxiliar2 = lobjstcCob.getRpt.CodCobert
        'Descrição
        lAuxiliar2 = lAuxiliar2 & "  " & lobjstcCob.getRpt.nomCobert

        lTamanho = Len(lAuxiliar2)
        If lTamanho < 70 Then
            lAuxiliar2 = lAuxiliar2 & Space(70 - lTamanho)
        Else
            lAuxiliar2 = Left(lAuxiliar2, 70)
        End If

        GeraCYX_PROP2 pLinhaCXY, "Cobertura " & lCountLine, pPosY, pPosX, False, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
        Print #pNumArqDados, lAuxiliar2

        'Valor da IS
        lAuxiliar3 = lobjstcCob.getRpt.limMaxGarantia
        lTamanho = Len(lAuxiliar3)
        If lTamanho < 14 Then
            lAuxiliar3 = Space(14 - lTamanho) & lAuxiliar3
        End If

        GeraCYX_PROP2 pLinhaCXY, "LMG COB " & lCountLine, pPosY, 12, False, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc, "COURIER NEW"
        Print #pNumArqDados, lAuxiliar3

        'f) Prêmio líquido .
        lAuxiliar3 = lobjstcCob.getRpt.premioLiq
        lTamanho = Len(lAuxiliar3)
        If lTamanho < 14 Then
            lAuxiliar3 = Space(14 - lTamanho) & lAuxiliar3
        End If

        GeraCYX_PROP2 pLinhaCXY, "PREMIO COB " & lCountLine, pPosY, 16.5, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc, "COURIER NEW"
        Print #pNumArqDados, lAuxiliar3
    Next lobjstcCob

    'Descricao do Premio Premio ("Prêmio NET R$ :" ou "Prêmio Líquido  R$ :")
    GeraCYX_PROP2 pLinhaCXY, "PREMIO LIQUIDO", pPosY, 13, False, "S", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
    Print #pNumArqDados, pobjstcPed.getRpt.DescricaoPremio

    'Coberturas contratadas - soma.
    GeraCYX_PROP2 pLinhaCXY, "PREMIO LIQUIDO", pPosY, 16.5, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc, "COURIER NEW"
    lAuxiliar2 = pobjstcPedLoc.getRpt.ValPrTotCober
    lTamanho = Len(lAuxiliar2)
    If lTamanho < 14 Then
        lAuxiliar2 = Space(14 - lTamanho) & lAuxiliar2
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '040. Informação cifrada .
    lAuxiliar2 = pobjstcPedLoc.getRpt.CodCoenferencia

    If pLinhaCXY = 42 Then pLinhaCXY = 41
    GeraCYX_PROP2 pLinhaCXY, "Linha", pPosY, pPosX, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    ' Tarifa 1056 e 1057 - Inicio
    If 0 = 1 Then    ' Desligado este bloco pois não vai mais entrar na tarifa
        '010. Coberturas contratadas
        lAuxiliar2 = String(75, "-") & " VALOR EM RISCO " & String(75, "-")
        Print #pNumArqDados, lAuxiliar2
        GeraCYX_PROP2 pLinhaCXY, "Titulo coberturas", pPosY, pPosX, True, "S", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc

        lAuxiliar2 = "COBERTURA" & String(130, " ") & "VALOR EM RISCO (R$)"
        Print #pNumArqDados, lAuxiliar2
        GeraCYX_PROP2 pLinhaCXY, "Titulo coberturas 2", pPosY, pPosX, True, "S", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
        lCountLine = 0
        For Each lobjstcCob In pobjstcPedLoc.ColecaoPedCobert
            'Código
            lCountLine = lCountLine + 1
            lAuxiliar2 = lobjstcCob.getRpt.CodCobert
            'Descrição
            lAuxiliar2 = lAuxiliar2 & "  " & lobjstcCob.getRpt.nomCobert

            lTamanho = Len(lAuxiliar2)
            If lTamanho < 70 Then
                lAuxiliar2 = lAuxiliar2 & Space(70 - lTamanho)
            Else
                lAuxiliar2 = Left(lAuxiliar2, 70)
            End If

            GeraCYX_PROP2 pLinhaCXY, "Cobertura " & lCountLine, pPosY, pPosX, False, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
            Print #pNumArqDados, lAuxiliar2

            'Valor da IS
            'lAuxiliar3 = lobjstcCob.getRpt.limMaxGarantia
            'lTamanho = Len(lAuxiliar3)
            'If lTamanho < 14 Then
            '    lAuxiliar3 = Space(14 - lTamanho) & lAuxiliar3
            'End If

            'GeraCYX_PROP2 pLinhaCXY, "LMG COB " & lCountLine, pPosY, 12, False, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc, "COURIER NEW"
            'Print #pNumArqDados, lAuxiliar3

            'f) Prêmio líquido .
            lAuxiliar3 = FormatNumber(lobjstcCob.Val_Risco_Cob, 2, vbTrue, vbUseDefault, vbTrue)
            lTamanho = Len(lAuxiliar3)
            If lTamanho < 18 Then
                lAuxiliar3 = Space(18 - lTamanho) & lAuxiliar3
            End If

            GeraCYX_PROP2 pLinhaCXY, "PREMIO COB " & lCountLine, pPosY, 15, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc, "COURIER NEW"
            Print #pNumArqDados, lAuxiliar3
        Next lobjstcCob
    End If
    ' Tarifa 1056 e 1057 - Fim

End Sub

Private Function mfImpCabec_MULTPRO2_Fonte12(ByRef pNumArqDados As Integer, ByRef pNumCXY As Integer, ByVal pobjstcPed As stcA46V702B, _
                                             pobjstcPedLoc As stcA46V708B, pNumDe As Integer, _
                                             pNumAte As Integer, pLinCXY As Long) As Boolean

'Função: cria arquivo de dados - proposta Anexo.

'Parâmetro de entrada...pNumArqDados...Número do arquivo de dados.


    Dim lAuxiliar1 As Double  'Auxiliar - utilizado apenas como exemplo.
    Dim lAuxiliar2 As String  'Auxiliar - utilizado apenas como exemplo.
    Dim lAuxiliar3 As String  'Auxiliar - utilizado apenas como exemplo.
    Dim lObjstcCorr As stcA46V705B    'Objeto stcA46V705B  (corretor).
    Dim lCountLine As Integer    'Contador de linhas.
    Dim lMensagem As String
    Dim lblnAchou As Boolean    'Indica se achou  o fator ou não.
    Dim lTemInfo As Boolean
    Dim lNumArquivo As Integer          'Número do arquivo.
    Dim lArqDados As String         'Endereço do arquivo de dados.

    mfImpCabec_MULTPRO2_Fonte12 = False
    lTemInfo = False

    Print #pNumCXY, "R Orientacao=Retrato; Papel=A4 210 x 297 mm; Altura=29.70; Largura=21.00; Vazio=0.00; ImgY=0.50; ImgX=0.50; ImgAltura=29.70; ImgLargura=21.00"

    '001. Título do seguro (de acordo com a base de dados).
    '     Neste caso deve ser colocado ANEXO DA PROPOSTA DE SEGURO YASUDA...
    '     RESIDENCIAL, CONDOMÍNIO ou EMPRESARIAL.

    Print #pNumArqDados, "Página: " & lQtdPaginas
    GeraCYX_PROP2 pLinCXY, "Página", 1, 18.5, False, "N", pNumCXY

    Print #pNumArqDados, UCase(pobjstcPed.gDesc_Ramo)
    GeraCYX_PROP2 pLinCXY, "Título", 1, 10.7, False, "N", pNumCXY
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
    GeraCYX_PROP2 pLinCXY, "Processo Susep", 1.4, 10.7, False, "N", pNumCXY
    '003. Emissão.
    lAuxiliar2 = Format$(Now, "dd/mm/yyyy hh:nn:ss")
    If pobjstcPed.Num_PI_Emissao <> 0 Then
        lAuxiliar2 = lAuxiliar2 & Space(20) & "PI: " & pobjstcPed.Num_PI_Emissao & "-" & pobjstcPed.Dig_PI_Emissao
    Else
        lAuxiliar2 = lAuxiliar2 & Space(20) & "PROPOSTA NÃO TRANSMITIDA"
    End If
    If pobjstcPed.Ind_Reentrada = 1 Then
        lAuxiliar2 = lAuxiliar2 & "**REENTRADA**"
    End If
    Print #pNumArqDados, lAuxiliar2
    GeraCYX_PROP2 pLinCXY, "Emissão", 1.9, 2.1, False, "N", pNumCXY

    '004. Siscota
    If pobjstcPed.objstcCtrlEmis.Tip_Cotac = 2 And pobjstcPed.objstcCtrlEmis.Num_Cotac <> 0 Then
        Print #pNumArqDados, "Siscota: " & pobjstcPed.objstcCtrlEmis.Num_Cotac
    Else
        Print #pNumArqDados, " "
    End If
    GeraCYX_PROP2 pLinCXY, "Siscota", 2.4, 12, False, "N", pNumCXY
    '005. Corretor (código e nome) .
    If pobjstcPed.ColecaoPedCorr.Count > 0 Then
        Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(1)
        lAuxiliar1 = lObjstcCorr.Cod_Corr
        lAuxiliar2 = lObjstcCorr.gNom_Corretor
    Else
        lAuxiliar1 = 0
        lAuxiliar2 = " "
    End If
    Print #pNumArqDados, Format$(lAuxiliar1, "0000") & " - " & lAuxiliar2
    GeraCYX_PROP2 pLinCXY, "Corretor", 3.1, 2, False, "N", pNumCXY

    '006. Número do pedido .
    lAuxiliar1 = pobjstcPed.Num_Ped
    Print #pNumArqDados, Format$(lAuxiliar1, "000000")
    GeraCYX_PROP2 pLinCXY, "Número do pedido", 3.5, 2.4, False, "N", pNumCXY

    '007. Número do item .
    lAuxiliar1 = pobjstcPedLoc.Num_Item
    Print #pNumArqDados, Format$(lAuxiliar1, "0000")
    GeraCYX_PROP2 pLinCXY, "Número do Item", 3.5, 6.4, False, "N", pNumCXY

    '008. Tarifa .
    lAuxiliar2 = pobjstcPed.Cod_Produto
    If pobjstcPed.Cod_Sequencia = 1 Then
        lAuxiliar2 = lAuxiliar2 & " (A)"
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    GeraCYX_PROP2 pLinCXY, "Número do Item", 3.5, 9, False, "N", pNumCXY
    '010. Nosso número .
    lAuxiliar2 = pobjstcPed.NOSSO_NUMERO
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    GeraCYX_PROP2 pLinCXY, "Número do Item", 3.5, 14.8, False, "N", pNumCXY

    '011. Dígito Verificador (Código Unicidade)
    Print #pNumArqDados, "DV:" & pobjstcPed.Cod_Unicidade & " - " & pobjstcPed.Cod_Ver_Unici
    GeraCYX_PROP2 pLinCXY, "Número do Item", 4.1, 14.8, False, "N", pNumCXY
    pLinCXY = 1
    mfImpCabec_MULTPRO2_Fonte12 = True

End Function
Private Function mfImpCabec_MULTCOTA_Fonte12(ByRef pNumArqDados As Integer, ByRef pNumCXY As Integer, ByVal pobjstcPed As stcA46V702B, _
                                             pobjstcPedLoc As stcA46V708B, pNumDe As Integer, _
                                             pNumAte As Integer, pLinCXY As Long) As Boolean

'Função: cria arquivo de dados - proposta Anexo.

'Parâmetro de entrada...pNumArqDados...Número do arquivo de dados.


    Dim lAuxiliar1 As Double  'Auxiliar - utilizado apenas como exemplo.
    Dim lAuxiliar2 As String  'Auxiliar - utilizado apenas como exemplo.
    Dim lAuxiliar3 As String  'Auxiliar - utilizado apenas como exemplo.
    Dim lObjstcCorr As stcA46V705B    'Objeto stcA46V705B  (corretor).
    Dim lCountLine As Integer    'Contador de linhas.
    Dim lMensagem As String
    Dim lblnAchou As Boolean    'Indica se achou  o fator ou não.
    Dim lTemInfo As Boolean
    Dim lNumArquivo As Integer          'Número do arquivo.
    Dim lArqDados As String         'Endereço do arquivo de dados.

    mfImpCabec_MULTCOTA_Fonte12 = False
    lTemInfo = False

    Print #pNumCXY, "R Orientacao=Retrato; Papel=A4 210 x 297 mm; Altura=29.70; Largura=21.00; Vazio=0.00; ImgY=0.50; ImgX=0.50; ImgAltura=29.70; ImgLargura=21.00"
    '001. Título do seguro (de acordo com a base de dados).
    '     Neste caso deve ser colocado ANEXO DA PROPOSTA DE SEGURO YASUDA...
    '     RESIDENCIAL, CONDOMÍNIO ou EMPRESARIAL.

    Print #pNumArqDados, "Página: " & lQtdPaginas
    GeraCYX_PROP2 pLinCXY, "Página", 1, 18.5, False, "N", pNumCXY

    Print #pNumArqDados, UCase(pobjstcPed.gDesc_Ramo)
    GeraCYX_PROP2 pLinCXY, "Título", 1, 10.3, False, "N", pNumCXY
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
    GeraCYX_PROP2 pLinCXY, "Processo Susep", 1.4, 10.7, False, "N", pNumCXY
    '003. Emissão.
    lAuxiliar2 = Format$(Now, "dd/mm/yyyy hh:nn:ss")
    If pobjstcPed.Ind_Reentrada = 1 Then
        lAuxiliar2 = lAuxiliar2 & "**REENTRADA**"
    End If
    Print #pNumArqDados, lAuxiliar2
    GeraCYX_PROP2 pLinCXY, "Emissão", 1.9, 2.1, False, "N", pNumCXY

    '004. Validade.
    If pobjstcPedLoc.Tip_Emissao = e_TipEmissao_RenYasuda Then
        If gM46V999.gCodProduto >= 978 And gM46V999.gCodProduto <= 1047 Then
            Print #pNumArqDados, Format$(CDate(gM46V999.gfForData(2, pobjstcPed.Dat_Pri_Calculo)) + 15, "dd/mm/yyyy")
        Else
            Print #pNumArqDados, Format$(CDate(gM46V999.gfForData(2, pobjstcPed.Dat_Pri_Calculo)) + 30, "dd/mm/yyyy")
        End If
    Else
        Print #pNumArqDados, Format$(CDate(gM46V999.gfForData(2, pobjstcPed.Dat_Pri_Calculo)) + 15, "dd/mm/yyyy")
    End If
    GeraCYX_PROP2 pLinCXY, "Validade", 1.9, 12, False, "N", pNumCXY

    '004. Siscota
    If pobjstcPed.objstcCtrlEmis.Tip_Cotac = 2 And pobjstcPed.objstcCtrlEmis.Num_Cotac <> 0 Then
        Print #pNumArqDados, "Siscota: " & pobjstcPed.objstcCtrlEmis.Num_Cotac
    Else
        Print #pNumArqDados, " "
    End If
    GeraCYX_PROP2 pLinCXY, "Siscota", 1.9, 14.8, False, "N", pNumCXY
    '005. Corretor (código e nome) .
    If pobjstcPed.ColecaoPedCorr.Count > 0 Then
        Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(1)
        lAuxiliar1 = lObjstcCorr.Cod_Corr
        lAuxiliar2 = lObjstcCorr.gNom_Corretor
    Else
        lAuxiliar1 = 0
        lAuxiliar2 = " "
    End If
    Print #pNumArqDados, Format$(lAuxiliar1, "0000") & " - " & lAuxiliar2
    GeraCYX_PROP2 pLinCXY, "Corretor", 2.7, 2, False, "N", pNumCXY

    '006. Número do pedido .
    lAuxiliar1 = pobjstcPed.Num_Ped
    Print #pNumArqDados, Format$(lAuxiliar1, "000000")
    GeraCYX_PROP2 pLinCXY, "Número do pedido", 3.1, 2.4, False, "N", pNumCXY

    '007. Número do item .
    lAuxiliar1 = pobjstcPedLoc.Num_Item
    Print #pNumArqDados, Format$(lAuxiliar1, "0000")
    GeraCYX_PROP2 pLinCXY, "Número do Item", 3.1, 6.4, False, "N", pNumCXY

    '008. Tarifa .
    lAuxiliar2 = pobjstcPed.Cod_Produto
    If pobjstcPed.Cod_Sequencia = 1 Then
        lAuxiliar2 = lAuxiliar2 & " (A)"
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    GeraCYX_PROP2 pLinCXY, "Número do Item", 3.1, 10, False, "N", pNumCXY

    '010. Nosso número .
    lAuxiliar2 = pobjstcPed.NOSSO_NUMERO
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    GeraCYX_PROP2 pLinCXY, "Número do Item", 3.1, 15.1, False, "N", pNumCXY

    pLinCXY = 1
    mfImpCabec_MULTCOTA_Fonte12 = True

End Function



'Subrotina nova para o Projeto Impressão em Fonte 12
Private Sub mpCriDados_DemonstrativoParcelamento_Coletiva(pNumArqDados As Integer, pLinhaCXY As Long, pPosY As Double, pPosX As Double, pNumArqCXY As Integer, pobjstcPed As stcA46V702B, pobjstcPedLoc As stcA46V708B)
    Dim lAuxiliar1 As Double
    Dim lAuxiliar2 As String
    Dim lAuxiliar3 As String
    Dim lTamanho As Integer
    Dim lValorParc As Variant
    Dim lMensagem As String
    Dim lI As Integer
    Dim lOrdem As Integer
    Dim lPriParcela As String
    Dim lDemParcelas As String    'Demais parcelas.
    Dim lobjstcJur As stcA46V711B
    Dim lJuros As String
    Dim lParcelas(12)

    If pLinhaCXY > 29 Then
        Call mpNovaPag(pLinhaCXY, pPosY, pPosX, pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc)
    End If


    'Calcula o fracionamento
    With pobjstcPed
        Call gM46V998.gfCalFracionamento(pobjstcPed, pobjstcPedLoc.Tip_Emissao, _
                                         lValorParc, lMensagem, gM46V111.gProJurIOF, .Cod_Reg, .Val_IOF, .Val_Parc_Min)
    End With

    '018. Formas de pagamento - primeira parcela a 30 dias.
    'Ficha 900665
    lOrdem = 1
    For lI = 1 To 12
        If lI > pobjstcPed.Qtd_Parc Then
            lAuxiliar2 = " "
            'Exit For
        Else
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
                If lTamanho < 9 Then
                    lDemParcelas = Space(9 - lTamanho) & lDemParcelas
                End If
                'd) Juros.
                Set lobjstcJur = New stcA46V711B
                On Error Resume Next
                Set lobjstcJur = pobjstcPed.ColecaoPedJuros.Item(pobjstcPed.NOSSO_NUMERO & Format(lI, "00"))
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
                lAuxiliar2 = Format(lOrdem, "00") & "    " & lPriParcela & "  " & lDemParcelas & "  " & lJuros
            Else
                lAuxiliar2 = "- "
            End If
            If lValorParc(lI, 3) <> 0 Then

                'b) Valor da primeira parcela.
                lAuxiliar1 = lValorParc(lI, 3)
                If lAuxiliar1 <> 0 Then
                    lPriParcela = Format$(lAuxiliar1, "standard")
                    lTamanho = Len(lPriParcela)
                    If lTamanho < 10 Then
                        lPriParcela = Space(10 - lTamanho) & lPriParcela
                    End If
                    'c) Valor das demais parcelas.
                    lAuxiliar1 = lValorParc(lI, 4)
                    lDemParcelas = Format$(lAuxiliar1, "standard")
                    lTamanho = Len(lDemParcelas)
                    If lTamanho < 9 Then
                        lDemParcelas = Space(9 - lTamanho) & lDemParcelas
                    End If
                    'd) Juros.
                    Set lobjstcJur = New stcA46V711B
                    On Error Resume Next
                    Set lobjstcJur = pobjstcPed.ColecaoPedJuros.Item(pobjstcPed.NOSSO_NUMERO & Format(lI, "00"))
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
                    lAuxiliar2 = lAuxiliar2 & "      " & lPriParcela & "  " & lDemParcelas & "  " & lJuros
                Else
                    lAuxiliar2 = lAuxiliar2 & " "
                End If
            Else
                lAuxiliar2 = lAuxiliar2 & " "
            End If
        End If
        If lAuxiliar2 <> "-  " Then
            Print #pNumArqDados, lAuxiliar2
            GeraCYX_PROP2 pLinhaCXY, "Parcela ", pPosY, 0.6, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc, "COURIER NEW"
        End If
    Next lI

End Sub


'Funcão nova para o Projeto Impressão em Fonte 12
Private Function mfCriDados_MULTCOLETIVA_Itens_Fonte12(ByVal pNumArqDados As Integer, _
                                                       ByVal pobjstcPed As stcA46V702B, _
                                                       ByVal pobjstcPedLoc As stcA46V708B, _
                                                       pNumDe As Integer, _
                                                       pNumAte As Integer) As Boolean
'Função: cria arquivo de dados - cotação coletiva - dados do item -

'Parâmetro de entrada...pNumArqDados...Número do arquivo de dados.

    Dim lAuxiliar1 As Double  'Auxiliar - utilizado apenas como exemplo.
    Dim lAuxiliar2 As String  'Auxiliar - utilizado apenas como exemplo.
    Dim lAuxiliar3 As String  'Auxiliar - utilizado apenas como exemplo.
    Dim lCodCobertura As Integer    'Código da cobertura.
    Dim lI As Integer    'Utilizado com For...Next.
    Dim lTamanho As Integer    'Tamanho.
    Dim lObjstcCorr As stcA46V705B
    Dim lobjstcCob As stcA46V704B    'Objeto da classe (stc) de cobertura.
    Dim lobjstcSeg As stcA46V716B    'Outra seguradora.
    Dim lobjstcBem As stcA46V715B    'Bens
    Dim lCountLine As Integer
    Dim lValTotCober As Double    'Soma das coberturas.
    Dim lDescFranquias() As String
    Dim lCount As Integer
    Dim lPrLiqItem As Double
    Dim lPrazo As Integer
    Dim lCodCobert As Integer
    Dim lNumMeses As Integer
    Dim lDescCom As Double
    Dim lDescExp As Double
    Dim lbdTab_Tipo_Fator As ADODB.Connection   'Arquivo: P00Tarifa.
    Dim lrsTab_Tipo_Fator As clsYasRecordSet    'Registro :Tab_Tipo_Fator.
    Dim lbdTbm_Regiao As ADODB.Connection
    Dim lrsTbm_Regiao As clsYasRecordSet
    Dim lobjstcFator As stcA46V720B       'Objeto da classe de estrutura stcA46V720B (TAB_PED_FAT_RISCO).
    Dim lblnAchou As Boolean        'Indica se achou  o fator ou não.
    Dim lMensagem As String
    Dim lDemParcelas As String    'Demais parcelas.
    Dim lQuaDias As Integer
    Dim lNumArqCXY As Integer
    Dim lPosX As Double
    Dim lPosY As Double
    Dim lLinCXY As Long
    Dim lValorComis As Single
    Dim lSelect As String

    mfCriDados_MULTCOLETIVA_Itens_Fonte12 = False

    'ROTINA CRIA CXY
    lCountPage = 1
    If gfIsMetaframe() Then
        mEndConfig = "L:\SYAS.CXY"

        Call gM46V999.gpGraLog(0, "Eh Metaframe forcou o L " & mEndConfig)
    Else
        mEndConfig = gM46V999.gAppPath & "SYAS.CXY"
    End If

    PodeGerarCYX = True
    lNumArqCXY = FreeFile

    lCountPage = 1
    lListaArq = mEndArq & "|" & mEndConfig & "|" & IIf(pobjstcPed.Cod_Tip_Doc = e_TipDoc_Cotacao, gM46V999.gEndImagCot3, gM46V999.gEndImagPro4) & ";"

    '    On Error Resume Next
    Open mEndConfig For Output Access Write Lock Read Write As #lNumArqCXY

    lPosX = 0.6
    lPosY = 11

    '''    If pobjstcPed.Cod_Tip_Doc = e_TipDoc_Cotacao Then
    Print #lNumArqCXY, "R Orientacao=Retrato; Papel=A4 210 x 297 mm; Altura=29.70; Largura=21.00; Vazio=0.00; ImgY=0.10; ImgX=0.10; ImgAltura=29.70; ImgLargura=21.00"
    '''    Else
    '''        Print #lNumArqCXY, "R Orientacao=Retrato; Papel=A4 210 x 297 mm; Altura=29.70; Largura=21.00; Vazio=0.00; ImgY=0.55; ImgX=0.55; ImgAltura=29.70; ImgLargura=21.00"
    '''    End If

    'Cabeçalho da pagina
    lQtdPaginas = lQtdPaginas + 1
    GeraCYX_PROP2 lLinCXY, "Página", 1, 18.5, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Título", 1.1, 10.7, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Processo Susep", 1.5, 10.7, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Emissão", 2, 2.2, False, "N", lNumArqCXY
    If pobjstcPed.Cod_Tip_Doc = e_TipDoc_Cotacao Then
        GeraCYX_PROP2 lLinCXY, "Validade", 2, 12, False, "N", lNumArqCXY
    End If
    GeraCYX_PROP2 lLinCXY, "Corretor", 2.82, 2.2, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Numero do pedido", 3.65, 0.71, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Número do item", 3.65, 5.61, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Tarifa", 3.65, 9, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Nosso numero", 3.65, 12.2, False, "N", lNumArqCXY
    If pobjstcPed.Cod_Tip_Doc = e_TipDoc_Proposta Then
        GeraCYX_PROP2 lLinCXY, "DV", 4.1, 14.8, False, "N", lNumArqCXY
    End If

    GeraCYX_PROP2 lLinCXY, "Dados do Risco - logradouro", 5.3, 0.71, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Dados do Risco - numero", 5.3, 10.5, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Dados do Risco - complemento", 5.3, 13.2, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Dados do Risco - bairro", 6.05, 0.71, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Dados do Risco - CEP", 6.05, 6, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Dados do Risco - cidade", 6.05, 8.6, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Dados do Risco - UF", 6.05, 13.9, False, "N", lNumArqCXY
    'Ficha 1000305
    GeraCYX_PROP2 lLinCXY, "Dados do seguro - construção", 6.05, 14.5, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Dados do seguro - bens compreendidos", 6.07, 17.1, False, "N", lNumArqCXY

    GeraCYX_PROP2 lLinCXY, "Dados do seguro - classificação", 6.9, 0.71, False, "N", lNumArqCXY
    'Ficha 1000305
    GeraCYX_PROP2 lLinCXY, "Dados do seguro - valor em risco", 6.9, 14.9, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Dados do seguro - Região", 7.7, 0.71, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Dados do seguro - vigência", 8.5, 0.71, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Dados do seguro - Ass24h", 8.5, 14.3, False, "N", lNumArqCXY

    '1000347 - Inclusão de informações da congênere em renovações
    If pobjstcPedLoc.Tip_Emissao = e_TipEmissao_RenCongenere Then
        GeraCYX_PROP2 lLinCXY, "Dados do seguro - tipo de emissão", 8.1, 11.5, False, "N", lNumArqCXY
        GeraCYX_PROP2 lLinCXY, "Congenere", 8.55, 9, False, "N", lNumArqCXY
        GeraCYX_PROP2 lLinCXY, "Apolice/Item", 9, 9, False, "N", lNumArqCXY
    Else
        GeraCYX_PROP2 lLinCXY, "Dados do seguro - tipo de emissão", 8.5, 9, False, "N", lNumArqCXY
    End If

    'Dados p/ Inspeção
    GeraCYX_PROP2 lLinCXY, "Dados para inspeção - nome da pessoa", 10.1, 5.55, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Dados para inspeção - telefone", 10.6, 4.35, False, "N", lNumArqCXY

    Print #pNumArqDados, "Página: " & lQtdPaginas

    '002. Título do seguro (de acordo com a base de dados).
    '     Neste caso deve ser colocado YASUDA RESIDENCIAL, YASUDA CONDOMÍNIO ou YASUDA EMPRESARIAL.
    Print #pNumArqDados, "YASUDA " & UCase(pobjstcPed.gDesc_Ramo)

    'Ficha 901223
    '001B. Código do Processo Susep.
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

    '''        '003. Número da página.
    '''        Print #pNumArqDados, "Página "; pNumDe & "/" & pNumAte
    '''        pNumDe = pNumDe + 1

    '004. Emissão.
    lAuxiliar2 = Format$(Now, "dd/mm/yyyy hh:nn:ss")
    If pobjstcPed.Num_PI_Emissao <> 0 Then
        lAuxiliar2 = lAuxiliar2 & Space(20) & "PI: " & pobjstcPed.Num_PI_Emissao & "-" & pobjstcPed.Dig_PI_Emissao
    Else
        If pobjstcPed.Cod_Tip_Doc = e_TipDoc_Proposta Then
            lAuxiliar2 = lAuxiliar2 & Space(20) & "PROPOSTA NÃO TRANSMITIDA"
        End If
    End If
    If pobjstcPed.Ind_Reentrada = 1 Then
        lAuxiliar2 = lAuxiliar2 & "**REENTRADA**"
    End If
    Print #pNumArqDados, lAuxiliar2


    '004. Validade.
    If pobjstcPed.Cod_Tip_Doc = e_TipDoc_Cotacao Then
        If pobjstcPedLoc.Tip_Emissao = e_TipEmissao_RenYasuda Then
            If gM46V999.gCodProduto >= 978 And gM46V999.gCodProduto <= 1047 Then
                Print #pNumArqDados, Format$(CDate(gM46V999.gfForData(2, pobjstcPed.Dat_Pri_Calculo)) + 15, "dd/mm/yyyy")
            Else
                Print #pNumArqDados, Format$(CDate(gM46V999.gfForData(2, pobjstcPed.Dat_Pri_Calculo)) + 30, "dd/mm/yyyy")
            End If
        Else
            Print #pNumArqDados, Format$(CDate(gM46V999.gfForData(2, pobjstcPed.Dat_Pri_Calculo)) + 15, "dd/mm/yyyy")
        End If
    End If

    '008. Corretor (código e nome) .
    If pobjstcPed.ColecaoPedCorr.Count > 0 Then
        Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(1)
        lAuxiliar1 = lObjstcCorr.Cod_Corr
        lAuxiliar2 = lObjstcCorr.gNom_Corretor
    Else
        lAuxiliar1 = 0
        lAuxiliar2 = " "
    End If
    Print #pNumArqDados, Format$(lAuxiliar1, "0000") & " - " & lAuxiliar2

    '006. Número do pedido .
    lAuxiliar1 = pobjstcPed.Num_Ped
    Print #pNumArqDados, Format$(lAuxiliar1, "000000")

    '007. Número do item .
    lAuxiliar1 = pobjstcPedLoc.Num_Item
    Print #pNumArqDados, Format$(lAuxiliar1, "0000")

    '009. Tarifa .
    lAuxiliar2 = pobjstcPed.Cod_Produto
    If pobjstcPed.Cod_Produto = 976 And pobjstcPed.Cod_Ramo = e_CodRamo_Empresarial Then
        lAuxiliar2 = lAuxiliar2 & " A"
    End If
    If pobjstcPed.Cod_Sequencia = 1 Then
        lAuxiliar2 = lAuxiliar2 & " (A)"
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '010. Nosso número .
    lAuxiliar2 = pobjstcPed.NOSSO_NUMERO
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    If pobjstcPed.Cod_Tip_Doc = e_TipDoc_Proposta Then
        '011. Dígito Verificador (Código Unicidade)
        Print #pNumArqDados, "DV:" & pobjstcPed.Cod_Unicidade & " - " & pobjstcPed.Cod_Ver_Unici
    End If

    '011. Dados do Risco - logradouro .
    lAuxiliar2 = pobjstcPedLoc.Cod_Tip_Logr_Risco & " " & pobjstcPedLoc.Nom_Logr_Risco
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '012. Dados do Risco - número .
    lAuxiliar2 = pobjstcPedLoc.objstcPedInsp.Num_Logr_Risco
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '013. Dados do Risco - complemento .
    lAuxiliar2 = pobjstcPedLoc.Dsc_Compl_Risco
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '014. Dados do Risco - bairro .
    lAuxiliar2 = pobjstcPedLoc.Nom_Bairro_Risco
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '017. Dados do Risco - CEP .
    lAuxiliar2 = pobjstcPedLoc.gCEP_Formatado
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '015. Dados do Risco - cidade .
    lAuxiliar2 = pobjstcPedLoc.Nom_Cid_Risco
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '016. Dados do Risco - UF .
    lAuxiliar2 = pobjstcPedLoc.Sig_UF_Risco
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '021. Dados do seguro - construção .
    'Ficha 1000305
    If gM46V999.gCod_Ramo = e_CodRamo_Empresarial Then
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

    'Ficha 1000305
    If gM46V999.gCod_Ramo = e_CodRamo_Empresarial Then
        Select Case pobjstcPedLoc.Tip_Cobert
        Case e_TipCobert_PredioeConteudo
            lAuxiliar2 = "   Prédio/Conteúdo"
        Case e_TipCobert_Predio
            lAuxiliar2 = "        Prédio    "
        Case e_TipCobert_Conteudo
            lAuxiliar2 = "       Conteúdo   "
        Case Else
            lAuxiliar2 = "   Prédio/Conteúdo"
        End Select
    Else
        lAuxiliar2 = "   Prédio/Conteúdo"
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '020. Dados do seguro - classificação .
    lAuxiliar2 = pobjstcPedLoc.Cod_Clasf & "-" & pobjstcPedLoc.Cod_Compl_Clasf & "- " & pobjstcPedLoc.Dsc_Clasf
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '024. Dados do seguro - valor em risco .
    lAuxiliar2 = "R$ " & Format(pobjstcPedLoc.Val_Vr, "standard")
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '022. Dados do seguro - região .
    'lAuxiliar2 = pobjstcPedloc.Cod_Regiao & " - " & Left(pobjstcPedloc.gDesc_Regiao, 38)
    'Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    If pobjstcPed.Cod_Produto <= gM46V999.TarifaAntiga Then
        lAuxiliar2 = pobjstcPedLoc.Cod_Regiao & " - " & Left(pobjstcPedLoc.gDesc_Regiao, 38)
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    Else
        'Nova Região
        If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTbm_Regiao, lMensagem) = False Then
            Call gM46V999.gpGraLog(2, lMensagem)
            Exit Function
        End If

        lSelect = "Select Cod_Faixa_Reg, DSC_REG from Tbm_Faixa_Regiao "
        lSelect = lSelect & " Where cod_Ramo = " & pobjstcPed.Cod_Ramo
        lSelect = lSelect & "    and Cod_Modalidade = " & pobjstcPedLoc.Cod_Produto
        lSelect = lSelect & "    and Cod_Tarif = " & pobjstcPed.Cod_Produto
        lSelect = lSelect & "    and COD_Sequencia <= " & pobjstcPed.Cod_Sequencia
        lSelect = lSelect & "    and Cep_Ini <= " & pobjstcPedLoc.Num_Cep_Risco
        lSelect = lSelect & "    and Cep_Fim >= " & pobjstcPedLoc.Num_Cep_Risco
        lSelect = lSelect & "    order by Cod_Sequencia desc "

        If gM46V999.gfObtRegistro(lbdTbm_Regiao, lSelect, lrsTbm_Regiao, lMensagem) = False Then
            Call gM46V999.gpGraLog(0, lMensagem)
            Call gM46V999.gpFecha2(lbdTbm_Regiao)
            Exit Function
        End If
        
        If Not lrsTbm_Regiao.EOF Then
            lAuxiliar2 = gM46V999.Check(lrsTbm_Regiao("Cod_Faixa_Reg")) '& " - " & Left(gM46V999.Check(lrsTbm_Regiao("DSC_REG")), 38)
        Else
            lAuxiliar2 = "Não cadastrada."
        End If

        lAuxiliar2 = lAuxiliar2 & "   -   PLANO: " & pobjstcPed.Desc_Plano & "   -   PRODUTO: " & Format(pobjstcPedLoc.Cod_Produto, "0000")
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    End If
   
    '023. Dados do seguro - vigência .
    lAuxiliar2 = "de " & gM46V999.gfForData(2, pobjstcPed.Dat_Ini_Vig) & _
               " a " & gM46V999.gfForData(2, pobjstcPed.Dat_Ter_Vig)
               
    'Se for perfil Calculo não será impresso conforme ficha 0701831
    If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda And _
       (gM46V999.PerfilAcesso = e_Perfil_CALC Or gM46V999.PerfilAcesso = e_Perfil_CALC_Gestor) Then
        lQuaDias = gM46V999.gfDifDatas(pobjstcPed.Dat_Ini_Vig, pobjstcPed.Dat_Ter_Vig)
        If lQuaDias >= 365 Then
            lAuxiliar2 = ""
        End If
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    
    Print #pNumArqDados, pobjstcPedLoc.getRpt.descricao_Ass24h

    '025. Dados do seguro - tipo de emissão .
    lAuxiliar2 = pobjstcPedLoc.gDesc_TipEmissao
    Select Case pobjstcPedLoc.Tip_Emissao
    Case e_TipEmissao_RenYasuda, e_TipEmissao_Endosso
        lAuxiliar2 = lAuxiliar2 & "  Apólice: " & pobjstcPedLoc.Num_Apol_Anterior & "  Item: " & pobjstcPedLoc.Num_Item_Anterior & "  Sinistro: " & IIf(pobjstcPedLoc.Ind_Sinistro = e_SimNao_Sim, "Sim", "Não")
    Case e_TipEmissao_RenCongenere
        '1000347 - Inclusão de informações da congênere em renovações
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", ":   " & Trim(lAuxiliar2))
        If pobjstcPedLoc.Ren_Con_Seguradora > 0 Then
            If Not gM46V999.gfNomSeguradora(pobjstcPedLoc.Ren_Con_Seguradora, lAuxiliar2, lMensagem) Then
                Call gM46V999.gpGraLog(1, lMensagem)
                Exit Function
            End If
        End If
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", Trim(lAuxiliar2))
        lAuxiliar2 = "Apólice: " & pobjstcPedLoc.Ren_Con_Apolice & "  Item: " & pobjstcPedLoc.Ren_Con_Item & "  Sinistro: " & IIf(pobjstcPedLoc.Ind_Sinistro = e_SimNao_Sim, "Sim", "Não")
    End Select
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)


    '018. Dados para inspeção - nome da pessoa para contato .
    lAuxiliar2 = pobjstcPedLoc.objstcPedInsp.Pess_Cont
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '019. Dados para inspeção - telefone para contato .
    lAuxiliar2 = pobjstcPedLoc.objstcPedInsp.Num_Tel
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    'Soma da Comissão para saber se o prêmio é NET
    If pobjstcPed.ColecaoPedCorr.Count > 0 Then
        Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(1)
        lValorComis = lObjstcCorr.Per_Corr
        If pobjstcPed.ColecaoPedCorr.Count > 1 Then
            Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(2)
            lValorComis = lValorComis + lObjstcCorr.Per_Corr
            If pobjstcPed.ColecaoPedCorr.Count > 2 Then
                Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(3)
                lValorComis = lValorComis + lObjstcCorr.Per_Corr
            End If
        End If
    End If

    'COBERTURAS
    lLinCXY = 9
    Call mpCriDados_CoberturasContratadas(pNumArqDados, lLinCXY, lPosY, lPosX, lNumArqCXY, pobjstcPed, pobjstcPedLoc)

    'PERÍODO INDENITÁRIO - somente imprimir caso seja contratado (informar a cobertura e o período contratado).
    Call mpCriDados_PeriodoIndenitario(pNumArqDados, lLinCXY, lPosY, lPosX, lNumArqCXY, pobjstcPed, pobjstcPedLoc)

    'FRANQUIAS - informar a cobertura e a respectiva franquia. Caso as coberturas contratadas não tenha franquia, informar texto "Não aplicável"
    Call mpCriDados_Franquia(pNumArqDados, lLinCXY, lPosY, lPosX, lNumArqCXY, pobjstcPed, pobjstcPedLoc)

    'QUESTIONÁRIO - imprimir os itens sempre, mesmo que não tenha sido selecionado nenhum item.
    Call mpCriDados_Questionario(pNumArqDados, lLinCXY, lPosY, lPosX, lNumArqCXY, pobjstcPed, pobjstcPedLoc)

    'RELAÇÃO DE BENS - informar os bens informados. Caso não tenha, informar texto "Não informado"
    Call mpCriDados_RelacaoBens(pNumArqDados, lLinCXY, lPosY, lPosX, lNumArqCXY, pobjstcPed, pobjstcPedLoc)

    'Relação de beneficiário
    Call mpCriDados_Beneficiario(pNumArqDados, lLinCXY, lPosY, lPosX, lNumArqCXY, pobjstcPed, pobjstcPedLoc)

    If gM46V999.gCodProduto >= 1014 Then
        Call mpImpClausulas(pNumArqDados, lLinCXY, lPosY, lPosX, lNumArqCXY, pobjstcPed, pobjstcPedLoc)
    End If

    If pobjstcPed.Cod_Tip_Doc = e_TipDoc_Proposta Then
        'DECLARAÇÕES
        Call mpCriDados_Declaracoes(pNumArqDados, lLinCXY, lPosY, lPosX, lNumArqCXY, pobjstcPed, pobjstcPedLoc)

        'LUCROS CESSANTES
        Call mpCriDados_LucrosCessantes(pNumArqDados, lLinCXY, lPosY, lPosX, lNumArqCXY, pobjstcPed, pobjstcPedLoc)
    End If

    'OBSERVAÇÕES - imprimir os textos fixos e texto digitado pelo analista no Syas.
    Call mpCriDados_Observacao(pNumArqDados, lLinCXY, lPosY, lPosX, lNumArqCXY, pobjstcPed, pobjstcPedLoc)

    Close #lNumArqCXY
    mfCriDados_MULTCOLETIVA_Itens_Fonte12 = True

End Function


'Funcão nova para o Projeto Impressão em Fonte 12
Private Function mfCriDados_MULTPRO3_Fonte12(ByVal pNumArqDados As Integer, _
                                             ByVal pobjstcPed As stcA46V702B, _
                                             ByVal pobjstcPedLoc As stcA46V708B, _
                                             ByVal pTotalPags As Integer) As Boolean
'Função: cria arquivo de dados - proposta coletiva - capa.

'Parâmetro de entrada...pNumArqDados...Número do arquivo de dados.

    Dim lAuxiliar1 As Double  'Auxiliar - utilizado apenas como exemplo.
    Dim lAuxiliar2 As String  'Auxiliar - utilizado apenas como exemplo.
    Dim lAuxiliar3 As String  'Auxiliar - utilizado apenas como exemplo.
    Dim lAuxiliar_01 As String
    Dim lAuxiliar_02 As String
    Dim lAuxiliar_03 As String

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
    Dim lobjstcPedJur As stcA46V711B
    Dim lNumArqCXY As Integer
    Dim lPosX As Double
    Dim lPosY As Double
    Dim lLinCXY As Long

    mfCriDados_MULTPRO3_Fonte12 = False


    'ROTINA CRIA CXY
    lCountPage = 1
    If gfIsMetaframe() Then
        mEndConfig = "L:\SYAS.CXY"

        Call gM46V999.gpGraLog(0, "Eh Metaframe forcou o L " & mEndConfig)
    Else
        mEndConfig = gM46V999.gAppPath & "SYAS.CXY"
    End If

    PodeGerarCYX = True
    lNumArqCXY = FreeFile

    lCountPage = 1
    lListaArq = mEndArq & "|" & mEndConfig & "|" & gM46V999.gEndImagPro3 & ";"

    On Error Resume Next
    Open mEndConfig For Output Access Write Lock Read Write As #lNumArqCXY

    lPosX = 0.6
    lPosY = 21

    Print #lNumArqCXY, "R Orientacao=Retrato; Papel=A4 210 x 297 mm; Altura=29.70; Largura=21.00; Vazio=0.00; ImgY=0.5; ImgX=0.5; ImgAltura=29.70; ImgLargura=21.00"

    'Cabeçalho da pagina
    lQtdPaginas = 1
    GeraCYX_PROP2 lLinCXY, "Página", 1, 18.5, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Título", 1, 10.7, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Processo Susep", 1.4, 10.7, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Emissão", 1.9, 2.1, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Corretor", 3.08, 2.1, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Referencia", 3.5, 2.6, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Tarifa", 3.5, 9.1, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Nosso Numero", 3.5, 14.8, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "DV", 4.1, 14.8, False, "N", lNumArqCXY

    'Proponente
    GeraCYX_PROP2 lLinCXY, "Nome do Segurado", 4.7, 1.8, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Código do Segurado", 4.7, 13.7, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Tipo de Pessoa", 4.7, 19, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "CPF/CNPJ", 5.5, 2.55, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Renda ou Faturamento", 6.05, 10, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Endereço", 6.75, 2.7, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Profissão ou Empresa", 6.05, 0.6, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Bairro", 7.55, 1.9, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Cidade", 7.55, 9.55, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "UF", 7.55, 15, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "CEP", 6.75, 12.9, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Telefone", 6.75, 17.5, False, "N", lNumArqCXY

    'Demostrativo do Premio
    GeraCYX_PROP2 lLinCXY, "Quantidade de Itens", 9.2, 4, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Premio Liquido Total", 10.5, 0.7, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Custo da Apolice", 10.5, 7, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "IOF%", 10, 11.6, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "IOF Valor", 10.5, 11.6, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Premio Total a Vista", 10.5, 16.4, False, "N", lNumArqCXY

    'Forma de Pagamento
    GeraCYX_PROP2 lLinCXY, "Número de Parcelas", 11.7, 3.8, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Parcelas Antecipadas", 12.1, 4.15, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Número do Bloqueto", 12.5, 4, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Juros A.m", 12.9, 2.65, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Juros de Mora", 11.7, 13.85, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Primeira Parcela", 12.1, 13.25, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Demais Parcela", 12.5, 13.25, False, "N", lNumArqCXY

    Print #pNumArqDados, "Página: " & lQtdPaginas

    '001. Título do documento.
    Print #pNumArqDados, UCase(pobjstcPed.gDesc_Ramo)

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
    'Print #pNumArqDados, "Emissão: " & Format$(Now, "dd/mm/yyyy hh:nn:ss")
    lAuxiliar2 = Format$(Now, "dd/mm/yyyy hh:nn:ss")
    If pobjstcPed.Num_PI_Emissao <> 0 Then
        lAuxiliar2 = lAuxiliar2 & Space(20) & "PI: " & pobjstcPed.Num_PI_Emissao & "-" & pobjstcPed.Dig_PI_Emissao
    Else
        lAuxiliar2 = lAuxiliar2 & Space(20) & "PROPOSTA NÃO TRANSMITIDA"
    End If
    If pobjstcPed.Ind_Reentrada = 1 Then
        lAuxiliar2 = lAuxiliar2 & "**REENTRADA**"
    End If
    Print #pNumArqDados, lAuxiliar2

    '''    '004. Número da página.
    '''    'Print #pNumArqDados, "Página 1/" & pTotalPags
    '''    Print #pNumArqDados, "1/" & pTotalPags

    '006. Corretor (código e nome) .
    If pobjstcPed.ColecaoPedCorr.Count > 0 Then
        Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(1)
        lAuxiliar1 = lObjstcCorr.Cod_Corr
        lAuxiliar2 = lObjstcCorr.gNom_Corretor
    Else
        lAuxiliar1 = 0
        lAuxiliar2 = " "
    End If
    Print #pNumArqDados, Format$(lAuxiliar1, "0000") & " - " & lAuxiliar2

    '005. Número do pedido .
    lAuxiliar1 = pobjstcPed.Num_Ped
    Print #pNumArqDados, Format$(lAuxiliar1, "000000")

    '007. Tarifa .
    lAuxiliar_02 = pobjstcPed.Cod_Produto
    If pobjstcPed.Cod_Produto = 976 And pobjstcPed.Cod_Ramo = e_CodRamo_Empresarial Then
        lAuxiliar2 = lAuxiliar2 & " A"
    End If
    If pobjstcPed.Cod_Sequencia = 1 Then
        lAuxiliar2 = lAuxiliar2 & " (A)"
    End If
    Print #pNumArqDados, lAuxiliar_02

    '008. Nosso número .
    Print #pNumArqDados, pobjstcPed.NOSSO_NUMERO

    '011. Dígito Verificador (Código Unicidade)
    Print #pNumArqDados, "DV:" & pobjstcPed.Cod_Unicidade & " - " & pobjstcPed.Cod_Ver_Unici

    '009. Proponente - nome do segurado .
    Print #pNumArqDados, IIf(pobjstcPed.Nom_Segurado = "", " ", pobjstcPed.Nom_Segurado)

    '010. Proponente - código do segurado - se houver .
    lAuxiliar1 = pobjstcPed.Cod_Segurado
    Print #pNumArqDados, Format$(lAuxiliar1, "0000000")

    '011. Proponente - tipo de pessoa .
    Select Case pobjstcPed.Tip_Pessoa
    Case e_TipPessoa_Juridica
        Print #pNumArqDados, "Jurídica"
        lAuxiliar2 = pobjstcPed.gCNPJ_CPF_Formatado & " - " & _
                     "Ramo de atividade: " & pobjstcPed.gDesc_RamoAtiv
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        If pobjstcPed.Val_RendaFatu > 0 Then
            lAuxiliar2 = "Faturamento  " & Format(pobjstcPed.Val_RendaFatu, "standard")
        Else
            lAuxiliar2 = "Faturamento não informado "
        End If
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    Case e_TipPessoa_Fisica
        Print #pNumArqDados, "Física"
        lAuxiliar2 = pobjstcPed.gCNPJ_CPF_Formatado & " - " & _
                     pobjstcPed.gCod_Doc_Segur & ": " & Trim(pobjstcPed.Num_Doc) & " - " & _
                     "Órgão emissor: " & Trim(pobjstcPed.Org_Exp) & " - " & _
                     "Emissão: " & gM46V999.gfForData(2, pobjstcPed.Dat_Emis_Doc)
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        If pobjstcPed.Val_RendaFatu > 0 Then
            lAuxiliar2 = "Renda  " & Format(pobjstcPed.Val_RendaFatu, "standard")
        Else
            lAuxiliar2 = "Renda não informada "
        End If
        'Ficha 1000307
        If gM46V999.gfPreenchido(pobjstcPed.objstcPedEndereco.gDesc_Reg) Then
            lAuxiliar2 = lAuxiliar2 & "      Cat.Registro: " & pobjstcPed.objstcPedEndereco.gDesc_Reg
            lAuxiliar2 = lAuxiliar2 & "     " & "N. Registro: " & pobjstcPed.objstcPedEndereco.Num_Reg_Cons
        End If
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    End Select

    '014. Proponente - endereço .
    lAuxiliar2 = pobjstcPed.objstcPedEndereco.Cod_Tip_Logr & " " & _
                 pobjstcPed.objstcPedEndereco.Nom_Logr & "," & _
                 pobjstcPed.objstcPedEndereco.Dsc_Compl_Logr
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    lAuxiliar2 = pobjstcPed.Nom_Empresa
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", "Profissão  " & lAuxiliar2)

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

    'Premio Liquido Total
    lAuxiliar1 = pobjstcPed.VAL_PR
    Print #pNumArqDados, Format$(lAuxiliar1, "standard")

    'Custo da Apolice
    lAuxiliar1 = pobjstcPed.Val_Cust_Emis
    Print #pNumArqDados, Format$(lAuxiliar1, "standard")

    'IOF
    Print #pNumArqDados, Format$(gM46V111.gProJurIOF, "standard") & "%"

    'IOF
    lAuxiliar1 = pobjstcPed.Val_IOF
    Print #pNumArqDados, Format$(lAuxiliar1, "standard")

    'Premio Total a Vista
    lAuxiliar1 = pobjstcPed.gValorTotal
    Print #pNumArqDados, Format$(lAuxiliar1, "standard")

    '025. Forma de pagamento: número de parcelas.
    Print #pNumArqDados, pobjstcPed.Qtd_Parc

    '026. Forma de pagamento: parcelas antecipadas.
    Print #pNumArqDados, IIf(pobjstcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec = e_SimNao_Sim, "1", "0")

    '027. Forma de pagamento: número do bloqueto.
    Print #pNumArqDados, pobjstcPed.objstcCtrlEmis.Num_Bloq & pobjstcPed.objstcCtrlEmis.DIG_BLOQUETO

    '028. Forma de pagamento: juros a. m.
    For Each lobjstcPedJur In pobjstcPed.ColecaoPedJuros
        If lobjstcPedJur.Num_Parcela = pobjstcPed.Qtd_Parc Then
            If pobjstcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec = e_SimNao_Sim Then
                pobjstcPed.objstcCtrlEmis.Per_Juros_Neg = lobjstcPedJur.Val_Ant
            Else
                pobjstcPed.objstcCtrlEmis.Per_Juros_Neg = lobjstcPedJur.Val_Pos
            End If
            Exit For
        End If
    Next lobjstcPedJur
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
        Print #pNumArqDados, "R$" & Format(lValorParc(pobjstcPed.Qtd_Parc, 1, e_FormaPagamento.e_FormaPagamento_Carne), "standard")
        pobjstcPed.Val_PriParcela = lValorParc(pobjstcPed.Qtd_Parc, 1, e_FormaPagamento.e_FormaPagamento_Carne)

        '031. Forma de pagamento: demais parcelas.
        Print #pNumArqDados, "R$" & Format(lValorParc(pobjstcPed.Qtd_Parc, 2, e_FormaPagamento.e_FormaPagamento_Carne), "standard")
        pobjstcPed.val_DemParcelas = lValorParc(pobjstcPed.Qtd_Parc, 2, e_FormaPagamento.e_FormaPagamento_Carne)
    Else
        '030. Forma de pagamento: primeira parcela.
        Print #pNumArqDados, "R$" & Format(lValorParc(pobjstcPed.Qtd_Parc, 1, e_FormaPagamento.e_FormaPagamento_Debito), "standard")
        pobjstcPed.Val_PriParcela = lValorParc(pobjstcPed.Qtd_Parc, 1, e_FormaPagamento.e_FormaPagamento_Debito)

        '031. Forma de pagamento: demais parcelas.
        Print #pNumArqDados, "R$" & Format(lValorParc(pobjstcPed.Qtd_Parc, 2, e_FormaPagamento.e_FormaPagamento_Debito), "standard")
        pobjstcPed.val_DemParcelas = lValorParc(pobjstcPed.Qtd_Parc, 2, e_FormaPagamento.e_FormaPagamento_Debito)
    End If


    lLinCXY = 1

    'Dados Débito em Conta
    Call mpCriDados_PagamentoDebito(pNumArqDados, lLinCXY, lPosY, lPosX, lNumArqCXY, pobjstcPed, pobjstcPedLoc)

    'COSSEGURO
    Call mpCriDados_Cosseguro(pNumArqDados, lLinCXY, lPosY, lPosX, lNumArqCXY, pobjstcPed, pobjstcPedLoc)

    'DADOS PARA CORRESPONDÊNCIA
    Call mpCriDados_EnderecoCobranca(pNumArqDados, lLinCXY, lPosY, lPosX, lNumArqCXY, pobjstcPed, pobjstcPedLoc)

    'USO DA SEGURADORA
    Call mpCriDados_UsoInterno(pNumArqDados, lLinCXY, lPosY, lPosX, lNumArqCXY, pobjstcPed, pobjstcPedLoc)


    'Local e data
    GeraCYX_PROP2 lLinCXY, "Local e data", 27, 0.5, False, "N", lNumArqCXY
    lAuxiliar2 = "São Paulo, " & Format$(Now, "dd") & " de " & Format$(Now, "mmmm") & " de " & Format$(Now, "yyyy")
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    'Corretor
    GeraCYX_PROP2 lLinCXY, "Corretor", 28, 13.5, False, "N", lNumArqCXY
    Print #pNumArqDados, lObjstcCorr.gNom_Corretor


    Close #lNumArqCXY
    mfCriDados_MULTPRO3_Fonte12 = True
End Function


'Funcão nova para o Projeto Impressão em Fonte 12
Private Function mfCriDados_MULTCOT2_Fonte12(ByVal pNumArqDados As Integer, ByVal pobjstcPed As stcA46V702B, _
                                             pobjstcPedLoc As stcA46V708B) As Boolean
'Função: cria arquivo de dados - cotação coletiva.

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
    Dim lPrazo As Integer
    Dim lCodCobert As Integer
    Dim lNumMeses As Integer
    Dim lDescCom As Double
    Dim lDescExp As Double
    Dim lValorComis As Single
    Dim lbdTbm_Regiao As ADODB.Connection
    Dim lrsTbm_Regiao As clsYasRecordSet
    Dim lQuaDias As Integer
    Dim lNumArqCXY As Integer
    Dim lPosX As Double
    Dim lPosY As Double
    Dim lLinCXY As Long


    'ROTINA CRIA CXY
    lCountPage = 1
    If gfIsMetaframe() Then
        mEndConfig = "L:\SYAS.CXY"

        Call gM46V999.gpGraLog(0, "Eh Metaframe forcou o L " & mEndConfig)
    Else
        mEndConfig = gM46V999.gAppPath & "SYAS.CXY"
    End If

    PodeGerarCYX = True
    lNumArqCXY = FreeFile

    '    On Error Resume Next
    Open mEndConfig For Output Access Write Lock Read Write As #lNumArqCXY

    Print #lNumArqCXY, "R Orientacao=Retrato; Papel=A4 210 x 297 mm; Altura=29.70; Largura=21.00; Vazio=0.00; ImgY=0.50; ImgX=0.50; ImgAltura=29.70; ImgLargura=21.00"

    'Cabeçalho da pagina
    lQtdPaginas = 1
    GeraCYX_PROP2 lLinCXY, "Página", 1, 18.5, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Título", 1, 10.3, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Processo Susep", 1.4, 10.7, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Emissão", 1.9, 2.1, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Validade", 1.85, 11.9, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Siscota", 2.4, 14.8, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Corretor", 2.68, 2, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Numero do pedido", 3.1, 2.5, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Tarifa", 3.1, 9.1, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Nosso numero", 3.1, 13, False, "N", lNumArqCXY

    'Dados do Proponente
    GeraCYX_PROP2 lLinCXY, "Proponente - nome", 5.1, 0.6, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Proponente - codigo", 5.1, 12.1, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Proponente - tipo de pessoa", 5.1, 16.5, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Proponente - CNPJ/CPF", 5.95, 0.6, False, "N", lNumArqCXY

    GeraCYX_PROP2 lLinCXY, "Quantidade de Itens", 7.2, 4, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Premio Liquido Total", 8, 0.7, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Custo da Apolice", 8, 7, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "IOF%", 7.55, 12.3, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "IOF Valor", 8, 12.6, False, "N", lNumArqCXY
    GeraCYX_PROP2 lLinCXY, "Premio Total a Vista", 8, 15.2, False, "N", lNumArqCXY

    mfCriDados_MULTCOT2_Fonte12 = False

    Print #pNumArqDados, "Página: " & lQtdPaginas

    '001. Título do seguro (de acordo com a base de dados).
    '     Neste caso deve ser colocado YASUDA RESIDENCIAL, YASUDA CONDOMÍNIO ou YASUDA EMPRESARIAL.
    lAuxiliar2 = " YASUDA " & UCase(pobjstcPed.gDesc_Ramo)

    Do While Len(lAuxiliar2) < 36
        lAuxiliar2 = " " & lAuxiliar2 & " "
    Loop
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '[FRH] - FICHA 901116
    '001B. Código do Processo Susep.
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
    '002. Processo Susep
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '003. Emissão.
    lAuxiliar2 = Format$(Now, "dd/mm/yyyy hh:nn:ss")
    If pobjstcPed.Ind_Reentrada = 1 Then
        lAuxiliar2 = lAuxiliar2 & "**REENTRADA**"
    End If
    Print #pNumArqDados, lAuxiliar2

    '004. Validade.
    If pobjstcPedLoc.Tip_Emissao = e_TipEmissao_RenYasuda Then
        If gM46V999.gCodProduto >= 978 And gM46V999.gCodProduto <= 1047 Then
            Print #pNumArqDados, Format$(CDate(gM46V999.gfForData(2, pobjstcPed.Dat_Pri_Calculo)) + 15, "dd/mm/yyyy")
        Else
            Print #pNumArqDados, Format$(CDate(gM46V999.gfForData(2, pobjstcPed.Dat_Pri_Calculo)) + 30, "dd/mm/yyyy")
        End If
    Else
        Print #pNumArqDados, Format$(CDate(gM46V999.gfForData(2, pobjstcPed.Dat_Pri_Calculo)) + 15, "dd/mm/yyyy")
    End If

    '005. Siscota
    If pobjstcPed.objstcCtrlEmis.Tip_Cotac = 2 And pobjstcPed.objstcCtrlEmis.Num_Cotac <> 0 Then
        Print #pNumArqDados, "Siscota: " & pobjstcPed.objstcCtrlEmis.Num_Cotac
    Else
        Print #pNumArqDados, " "
    End If

    '006. Corretor (código e nome) .
    If pobjstcPed.ColecaoPedCorr.Count > 0 Then
        Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(1)
        lAuxiliar1 = lObjstcCorr.Cod_Corr
        lAuxiliar2 = lObjstcCorr.gNom_Corretor
    Else
        lAuxiliar1 = 0
        lAuxiliar2 = " "
    End If

    Print #pNumArqDados, Format$(lAuxiliar1, "0000") & " - " & lAuxiliar2

    '007. Número do pedido .
    lAuxiliar1 = pobjstcPed.Num_Ped
    Print #pNumArqDados, Format$(lAuxiliar1, "000000")

    '009. Tarifa .
    lAuxiliar2 = pobjstcPed.Cod_Produto
    If pobjstcPed.Cod_Sequencia = 1 Then
        lAuxiliar2 = lAuxiliar2 & " (A)"
    End If
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '010. Nosso número .
    lAuxiliar2 = pobjstcPed.NOSSO_NUMERO
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '011. Proponente - nome do segurado .
    lAuxiliar2 = pobjstcPed.Nom_Segurado
    Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)

    '012. Proponente - código do segurado - se houver .
    lAuxiliar1 = pobjstcPed.Cod_Segurado
    Print #pNumArqDados, Format$(lAuxiliar1, "0000000")

    '013. Proponente - tipo de pessoa .
    Select Case pobjstcPed.Tip_Pessoa
    Case e_TipPessoa_Juridica
        Print #pNumArqDados, "Jurídica"
        '013 e 014. Proponente - CNPJ ou CPF .
        'lAuxiliar2 = "CNPJ.........:"
        lAuxiliar2 = "              "
        lAuxiliar2 = pobjstcPed.gCNPJ_CPF_Formatado
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        '''            '015. Identificação
        '''            lAuxiliar2 = "Ramo de atividade: " & pobjstcPed.gDesc_RamoAtiv
        '''            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    Case e_TipPessoa_Fisica
        Print #pNumArqDados, "Física"
        '013 e 014. Proponente - CNPJ ou CPF .
        'lAuxiliar2 = "CPF...........:"
        lAuxiliar2 = "              "
        lAuxiliar2 = pobjstcPed.gCNPJ_CPF_Formatado
        'Ficha 1000307
        If gM46V999.gfPreenchido(pobjstcPed.objstcPedEndereco.gDesc_Reg) Then
            lAuxiliar2 = lAuxiliar2 & "       Cat.Registro: " & pobjstcPed.objstcPedEndereco.gDesc_Reg
            lAuxiliar2 = lAuxiliar2 & "        " & "Número Registro: " & pobjstcPed.objstcPedEndereco.Num_Reg_Cons
        End If
        Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
        '''            '015. Identificação
        '''            lAuxiliar2 = pobjstcPed.gCod_Doc_Segur & ": " & pobjstcPed.Num_Doc & "   " & _
         '''                   "Órgão emissor: " & pobjstcPed.Org_Exp & "   " & _
         '''                   "Emissão: " & gM46V999.gfForData(2, pobjstcPed.Dat_Emis_Doc)
        '''            Print #pNumArqDados, IIf(lAuxiliar2 = "", " ", lAuxiliar2)
    End Select

    'Quantidade de Itens
    lAuxiliar1 = pobjstcPed.ColecaoPedLoc.Count
    Print #pNumArqDados, lAuxiliar1

    'Soma da Comissão para saber se o prêmio é NET
    If pobjstcPed.ColecaoPedCorr.Count > 0 Then
        Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(1)
        lValorComis = lObjstcCorr.Per_Corr
        If pobjstcPed.ColecaoPedCorr.Count > 1 Then
            Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(2)
            lValorComis = lValorComis + lObjstcCorr.Per_Corr
            If pobjstcPed.ColecaoPedCorr.Count > 2 Then
                Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(3)
                lValorComis = lValorComis + lObjstcCorr.Per_Corr
            End If
        End If
    End If

    'Premio Liquido Total
    lAuxiliar1 = pobjstcPed.VAL_PR
    lAuxiliar2 = "Prêmio Líquido Total: "
    If lValorComis = 0 Then
        lAuxiliar2 = "Prêmio NET total: "
    End If
    If pobjstcPed.Cod_Reg > 100 Then
        If pobjstcPed.VAL_PR = gM46V999.gProPreMinEndosso Then
            lAuxiliar2 = "Prêmio mínimo total: "
        End If
    End If
    Print #pNumArqDados, lAuxiliar2 & Format$(lAuxiliar1, "standard")

    'Custo da Apolice
    lAuxiliar1 = pobjstcPed.Val_Cust_Emis
    Print #pNumArqDados, IIf(lValorComis > 0, "Custo da Apólice: " & Format$(lAuxiliar1, "standard"), " ")

    'IOF %
    Print #pNumArqDados, IIf(lValorComis > 0, "IOF( " & Format$(gM46V111.gProJurIOF, "standard") & "% )", " ")

    'IOF R$
    lAuxiliar1 = pobjstcPed.Val_IOF
    Print #pNumArqDados, IIf(lValorComis > 0, Format$(lAuxiliar1, "standard"), " ")

    'Premio Total a Vista
    lAuxiliar1 = pobjstcPed.gValorTotal
    Print #pNumArqDados, IIf(lValorComis > 0, "Prêmio Total a Vista: " & Format$(lAuxiliar1, "standard"), " ")

    If lValorComis > 0 Then
        lPosX = 0.6
        lPosY = 10.1
        'Formas de Pagamento
        Call mpCriDados_DemonstrativoParcelamento_Coletiva(pNumArqDados, lLinCXY, lPosY, lPosX, lNumArqCXY, pobjstcPed, pobjstcPedLoc)
    Else
        lPosX = 0.6
        lPosY = 16.1
    End If

    lCountPage = 1
    lListaArq = mEndArq & "|" & mEndConfig & "|" & gM46V999.gEndImagCot2 & ";"

    Call mpCriDados_UsoInterno(pNumArqDados, lLinCXY, lPosY, lPosX, lNumArqCXY, pobjstcPed, pobjstcPedLoc)

    Close #lNumArqCXY
    mfCriDados_MULTCOT2_Fonte12 = True
End Function


Private Sub mpQuebraTextoFranquia(ByRef pTextos() As String, ByVal pTamanho As Integer)
    Dim intCount As Integer
    Dim strTexto As String
    Dim lPosicao As Integer
    Dim lTexto() As String

    strTexto = ""
    ReDim lTexto(0)
    For intCount = 0 To UBound(pTextos)
        strTexto = pTextos(intCount)
        If strTexto = "" Then Exit For
        Do
            'encontra primeiro espaço antes do tamanho determinado
            lPosicao = InStrRev(strTexto, " ", pTamanho, vbTextCompare)
            If lPosicao = 0 Then
                lTexto(UBound(lTexto)) = strTexto
                ReDim Preserve lTexto(UBound(lTexto) + 1)
                Exit Do
            End If
            'guarda texto até o primeiro espaço antes do tamanho da linha definido
            lTexto(UBound(lTexto)) = Mid(strTexto, 1, lPosicao)
            'retira da variável o texto que já foi separado acima
            strTexto = Mid(strTexto, lPosicao, Len(strTexto))
            strTexto = "   " & strTexto
            ReDim Preserve lTexto(UBound(lTexto) + 1)
        Loop Until lPosicao = 0
    Next

    'redimenciona o array p/ ficar igual ao de entrada
    If UBound(lTexto) < UBound(pTextos) Then
        ReDim Preserve lTexto(UBound(lTexto) + (UBound(pTextos) - UBound(lTexto)))
    End If

    'atualiza o parametro de array entrada com os dados formatados
    pTextos = lTexto

End Sub


Private Sub mpCriDados_PagamentoDebito(pNumArqDados As Integer, pLinhaCXY As Long, pPosY As Double, pPosX As Double, pNumArqCXY As Integer, pobjstcPed As stcA46V702B, pobjstcPedLoc As stcA46V708B)
    Dim lAuxiliar As String

    If pLinhaCXY > 40 Then
        Call mpNovaPag(pLinhaCXY, pPosY, pPosX, pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc)
    End If

    If pobjstcPed.Num_Cta_Bas_Cor > 0 Then
        lAuxiliar = String(72, "-") & " DADOS DO PAGAMENTO " & String(72, "-")
        Print #pNumArqDados, lAuxiliar
        GeraCYX_PROP2 pLinhaCXY, "Título Dados do Pagamento", pPosY, pPosX, True, "S", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc

        '3. Banco...
        GeraCYX_PROP2 pLinhaCXY, "Banco", pPosY, 1, False, "S", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
        Print #pNumArqDados, "Banco:"

        GeraCYX_PROP2 pLinhaCXY, "Nome Banco", pPosY, 2.3, False, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
''20170123: Ajustado para buscar descricao dos bancos de debito do banco de dados (TBM_BANCOS)
''pobjstcPed.gDesc_Banco passa a ser uma property q busca a descricao no banco.
'        Select Case pobjstcPed.Cod_Bco_Cobr
'        Case 347, 356  'Sudameris.
'            Print #pNumArqDados, "Real"
'        Case 479    'Bank Boston.
'            Print #pNumArqDados, "Bank Boston"
'        Case 237    'Bradesco.
'            Print #pNumArqDados, "Bradesco"
'        Case 341    'Banco Itaú.
'            Print #pNumArqDados, "Itaú"
'        Case 409    'Unibanco.
'            Print #pNumArqDados, "Unibanco"
'        Case 399    'HSBC
'            Print #pNumArqDados, "HSBC"
'        Case 33    'Santander
'            Print #pNumArqDados, "Santander"
'        Case 47    'Banese
'            Print #pNumArqDados, "Banese"
'        Case Else
'            Print #pNumArqDados, " -"
'        End Select

        Print #pNumArqDados, pobjstcPed.gDesc_Banco
        
        'AGEN - Agência.
        GeraCYX_PROP2 pLinhaCXY, "Agencia", pPosY, 4.7, False, "S", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
        Print #pNumArqDados, "Agência:"
        GeraCYX_PROP2 pLinhaCXY, "Dados da Agencia", pPosY, 6.2, False, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
        Print #pNumArqDados, pobjstcPed.gAgenciaDig

        'CCOR - Conta corrente.
        GeraCYX_PROP2 pLinhaCXY, "Conta Corrente", pPosY, 8, False, "S", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
        Print #pNumArqDados, "Conta Corrente:"

        GeraCYX_PROP2 pLinhaCXY, "Dados Conta Corrente", pPosY, 10.7, True, "N", pNumArqCXY, pNumArqDados, pobjstcPed, pobjstcPedLoc
        lAuxiliar = ""
        Call mpForConta1(pobjstcPed.Cod_Bco_Cobr, pobjstcPed.mNum_Cta_Bas_Cor_Compl, lAuxiliar)
        Print #pNumArqDados, lAuxiliar

        pPosY = pPosY + 0.2

    End If
End Sub


Private Sub mpQuebraTextoArray(ByRef pTextos() As String, ByVal pTamanho As Integer)
    Dim intCount As Integer
    Dim strTexto As String
    Dim lPosicao As Integer
    Dim lTexto() As String


    strTexto = ""
    For intCount = 0 To UBound(pTextos)
        strTexto = strTexto & pTextos(intCount)
    Next

    ReDim lTexto(0)
    ''    For intCount = 0 To UBound(pTextos)
    ''        strTexto = pTextos(intCount)
    If strTexto = "" Then Exit Sub
    Do
        'encontra primeiro espaço antes do tamanho determinado
        lPosicao = InStrRev(strTexto, " ", pTamanho, vbTextCompare)
        If lPosicao = 0 Then
            lTexto(UBound(lTexto)) = strTexto
            ReDim Preserve lTexto(UBound(lTexto) + 1)
            Exit Do
        End If
        'guarda texto até o primeiro espaço antes do tamanho da linha definido
        lTexto(UBound(lTexto)) = Mid(strTexto, 1, lPosicao)
        'retira da variável o texto que já foi separado acima
        strTexto = Mid(strTexto, lPosicao + 1, Len(strTexto))
        '''            strTexto = "   " & strTexto
        ReDim Preserve lTexto(UBound(lTexto) + 1)
    Loop Until lPosicao = 0
    '''    Next

    'redimenciona o array p/ ficar igual ao de entrada
    If UBound(lTexto) < UBound(pTextos) Then
        ReDim Preserve lTexto(UBound(lTexto) + (UBound(pTextos) - UBound(lTexto)))
    End If

    'atualiza o parametro de array entrada com os dados formatados
    pTextos = lTexto

End Sub

'Ficha 1000284
Public Sub mpBuscaLimFranquia_CobertDesativado(ByVal pCodCobert As Integer, ByRef ValFraqMin As Double)
'Desativado Simony
    Dim lbdTbm_Fat_Franquia As ADODB.Connection         'Arquivo: P00Multi.
    Dim lMensagem As String         'Mensagem.
    Dim lrsTbm_Fat_Franquia As clsYasRecordSet          'Registro : Tbm_Fator_Franquia.
    Dim lCodSequencia As Integer
    Dim lCod_Clasf As String
    Dim lCod_Compl_Clasf As String
    Dim lCod_Constr As Integer
    Dim lTip_Cobert As Integer
    Dim lCEP As String
    Dim lSelect As String
    'Ficha 1000305
    If gM46V999.gCod_Ramo = e_CodRamo_Empresarial Then
        Select Case frmT46V101A.cboConstrucao.ItemData(frmT46V101A.cboConstrucao.ListIndex)
        Case 1
            lCod_Constr = e_CodContr_Superior
        Case 2
            lCod_Constr = e_CodContr_Inferior
        Case 3
            lCod_Constr = e_CodContr_Mista
        Case Else
            lCod_Constr = e_CodContr_Superior
        End Select
    Else
        If frmT46V101A.optContrucao(1).Value = True Then
            lCod_Constr = e_CodContr_Inferior
        Else
            lCod_Constr = e_CodContr_Superior
        End If
    End If

    'Ficha 1000305
    '== Tipo Cobertura
    If gM46V999.gCod_Ramo = e_CodRamo_Empresarial Then
        Select Case frmT46V101A.cboTipCobert.ItemData(frmT46V101A.cboTipCobert.ListIndex)
        Case 0
            lTip_Cobert = e_TipCobert_PredioeConteudo
        Case 1
            lTip_Cobert = e_TipCobert_Predio
        Case 2
            lTip_Cobert = e_TipCobert_Conteudo
        Case Else
            lTip_Cobert = e_TipCobert_PredioeConteudo
        End Select
    Else
        lTip_Cobert = 0
    End If
    gM46V999.gTipCobertura = lTip_Cobert

    lCod_Clasf = Left(frmT46V101A.txtClassificacao.Tag, 4)
    lCod_Compl_Clasf = Mid(frmT46V101A.txtClassificacao.Tag, 5, 2)

    If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTbm_Fat_Franquia, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If

    lCodSequencia = gM46V999.gCodSequencia

NovaSequencia1:

    '1 - Procura Registro de exceção com a chave completa
    'Procura o valor do Val_Franq_Min e Perc_Franq
    lCEP = val(Trim(Replace(frmT46V101A.txtRisCEP.Text, "-", "")))

    lSelect = " Select Val_Franq_Min as Min from Tbm_Fator_Franquia "
    lSelect = lSelect & " Where Cod_Ramo = " & gM46V999.gCod_Ramo
    lSelect = lSelect & " and Cod_Modalidade = " & frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex)
    lSelect = lSelect & " and Cod_Tip_Reg = 'E' "
    If gM46V999.gCodProduto >= 940 Then
        lSelect = lSelect & " and ((Cod_Clasf = '" & lCod_Clasf & "'"
        lSelect = lSelect & " and Cod_Compl_Clasf = '" & lCod_Compl_Clasf & "')"
        lSelect = lSelect & " or (Cod_Clasf = '0'))"
    Else
        lSelect = lSelect & " and Cod_Clasf = '" & lCod_Clasf & "'"
        lSelect = lSelect & " and Cod_Compl_Clasf = '" & lCod_Compl_Clasf & "'"
    End If
    lSelect = lSelect & " and Cod_Cobert = " & pCodCobert
    lSelect = lSelect & " and Cod_Construcao in (0 , " & lCod_Constr & ")"
    lSelect = lSelect & " and Cod_Tarif = " & gM46V999.gCodProduto
    lSelect = lSelect & " and Cod_Sequencia <= " & lCodSequencia

    If gM46V999.gCodProduto >= 940 Then
        lSelect = lSelect & " AND ((CEP_INICIO <= " & lCEP
        lSelect = lSelect & " AND CEP_FIM >= " & lCEP & " ) "
        lSelect = lSelect & " OR (CEP_INICIO =  0 ))"
        lSelect = lSelect & " order by Prioridade desc "
    End If

    If gM46V999.gfObtRegistro(lbdTbm_Fat_Franquia, lSelect, lrsTbm_Fat_Franquia, lMensagem) = False Then
        Call gM46V999.gpFecha2(lbdTbm_Fat_Franquia)
        Call gM46V999.gpGraLog(2, lMensagem)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    '2 - Se não encontrar a chave completa vai procurar o padrão
    If lrsTbm_Fat_Franquia.EOF Then
        If lCodSequencia > 0 Then
            lCodSequencia = lCodSequencia - 1
            GoTo NovaSequencia1
        Else
            GoTo ProcuraDefault
        End If
    End If

    If lrsTbm_Fat_Franquia("Min") = 0 Then
        GoTo ProcuraDefault
    End If

    ValFraqMin = lrsTbm_Fat_Franquia("Min")

    GoTo FimBusca

ProcuraDefault:


    '3 - Procura o registro somente com Ramo, Modalidade, Cobertura e TipoAtividade
    'Procura o valor do Val_Franq_Min, Perc_Franq

    lCodSequencia = gM46V999.gCodSequencia

NovaSequencia2:

    lSelect = " Select Val_Franq_Min as Min from Tbm_Fator_Franquia "
    lSelect = lSelect & " Where Cod_Ramo = " & gM46V999.gCod_Ramo
    lSelect = lSelect & " and Cod_Modalidade = " & frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex)
    lSelect = lSelect & " and Cod_Tip_Reg = 'D' "
    lSelect = lSelect & " and Cod_Cobert = " & pCodCobert
    lSelect = lSelect & " and Tip_Atividade in (3," & gTip_Atividade & ")"
    lSelect = lSelect & " and Cod_Tarif = " & gM46V999.gCodProduto
    lSelect = lSelect & " and Cod_Sequencia <= " & lCodSequencia

    If gM46V999.gfObtRegistro(lbdTbm_Fat_Franquia, lSelect, lrsTbm_Fat_Franquia, lMensagem) = False Then
        Call gM46V999.gpFecha2(lbdTbm_Fat_Franquia)
        Call gM46V999.gpGraLog(2, lMensagem)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    If lrsTbm_Fat_Franquia.EOF Then
        If lCodSequencia > 0 Then
            lCodSequencia = lCodSequencia - 1
            GoTo NovaSequencia2
        Else
            GoTo FimBusca
        End If
    End If

    If lrsTbm_Fat_Franquia("Min") = 0 Then
        GoTo FimBusca
    End If

    ValFraqMin = lrsTbm_Fat_Franquia("Min")

FimBusca:
    Call gM46V999.gpFecha2(lbdTbm_Fat_Franquia)
    '1001135 - Corretor Volks - Franquia Especial Concessionaria
    If gM46V999.gCodProduto >= 1002 Then
        If (InStr(1, "00008001/00008002/00008003/00008004/00008005/00008006/00008008/00008011/00008014/00008028", Format(frmT46V102A.txtCorretor(1).Tag, "00000000"), vbTextCompare) > 0) And _
           (val(lCod_Clasf) = 22 And val(lCod_Compl_Clasf) = 13) And _
           (InStr(1, "0038/0064/0066", Format(pCodCobert, "0000"), vbTextCompare) > 0) Then
            ValFraqMin = 2000
        End If
    End If
End Sub


'Ficha 1000655
