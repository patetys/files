VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmT46V163A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impressão das apólices"
   ClientHeight    =   4290
   ClientLeft      =   2835
   ClientTop       =   1410
   ClientWidth     =   5115
   Icon            =   "T46V163A.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstLista 
      Height          =   2310
      Left            =   0
      Style           =   1  'Checkbox
      TabIndex        =   6
      Top             =   1560
      Width           =   5055
   End
   Begin VB.CheckBox chkSel 
      Caption         =   "Seleciona todos"
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   3960
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin Threed.SSPanel panAviso 
      Height          =   1425
      Left            =   0
      TabIndex        =   2
      Top             =   80
      Visible         =   0   'False
      Width           =   5100
      _Version        =   65536
      _ExtentX        =   8996
      _ExtentY        =   2514
      _StockProps     =   15
      Caption         =   "Aviso..."
      ForeColor       =   12582912
      BackColor       =   12632319
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox RchTemp 
      Height          =   375
      Left            =   3915
      TabIndex        =   1
      Top             =   1620
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393217
      TextRTF         =   $"T46V163A.frx":030A
   End
   Begin RichTextLib.RichTextBox RchFinal 
      Height          =   375
      Left            =   180
      TabIndex        =   0
      Top             =   1575
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393217
      TextRTF         =   $"T46V163A.frx":0390
   End
   Begin Threed.SSCommand cmdOk 
      Height          =   300
      Left            =   2760
      TabIndex        =   3
      Top             =   3960
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1764
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Ok"
      ForeColor       =   12582912
   End
   Begin Threed.SSCommand cmdCancelar 
      Cancel          =   -1  'True
      Height          =   300
      Left            =   3960
      TabIndex        =   4
      Top             =   3960
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1764
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Cancelar"
      ForeColor       =   12582912
   End
   Begin VB.DriveListBox Drive 
      Height          =   315
      Left            =   2040
      TabIndex        =   7
      Top             =   2880
      Width           =   375
   End
End
Attribute VB_Name = "frmT46V163A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim mTipoDocumento As e_TipDoc
Dim lConfig As Boolean
Dim mobjPed As clsA46V702A
Dim mContadorPagina As Integer
Dim mLetraMapeamento As String
Dim mErro As Integer

Private Sub mpRetornaLetraParaMapear()
    Dim lI As Long
    Dim lAsc As Integer
    Dim lExiste As Boolean
    For lAsc = 102 To 122
        mLetraMapeamento = Chr(lAsc)
        lExiste = False
        Call gM46V999.gpGraLog(0, " Verificando letra " & mLetraMapeamento)
        For lI = 0 To Drive.ListCount - 1
            If InStr(1, Drive.List(lI), mLetraMapeamento & ":") Then
                If Drive.List(lI) = mLetraMapeamento & ": [\\SPX20423PSLUYNL\dados]" Then
                    Exit For
                End If
                Call gM46V999.gpGraLog(0, " Letra " & mLetraMapeamento & " já mapeada -->" & Drive.List(lI))
                lExiste = True
                Exit For
            End If
        Next lI
        If lExiste = False Then
            If gfGeraBatNetUse(True) Then Exit Sub
        End If
    Next lAsc
    mLetraMapeamento = ""
    panAviso.Caption = "Não foi localizada nenhuma letra para mapear o diretório necessário para esta emissão." & vbCrLf & "Por favor desconectar alguma unidade e tentar novamente este procedimento."
    Call gM46V999.gpGraLog(1, panAviso.Caption)
End Sub
Public Function gfGeraBatNetUse(ByVal lOpcao As Boolean) As Boolean
'Procedure utilizada para criar um arquivo .Bat contendo
'instruções para mapear uma determinada unidade de rede, temporariamente
'Parametro de entrada..... lOpcao = TRUE  >> Criar mapeamento
'                                   FALSE >> Desconetar mapeamento
    Dim lEndArqBAT As String
    Dim lNumArqBAT As String

    On Error GoTo ControleErro
    '1- exclui o ultimo arquivo se existir
    lEndArqBAT = gM46V999.gAppPath & "P46V001_NETUSE.BAT"
    If gM46V999.gfPreenchido(Dir(lEndArqBAT)) = True Then
        Kill lEndArqBAT
    End If

    '2 - Inicia o a criação do arquivo
    lNumArqBAT = FreeFile

    Open lEndArqBAT For Output Access Write As #lNumArqBAT

    Print #lNumArqBAT, "@ECHO OFF"
    Print #lNumArqBAT, "@CLS"

    If lOpcao = True Then
        Print #lNumArqBAT, "net use " & mLetraMapeamento & ": \\SPX20423PSLUYNL\dados teste /user:transfer"
    Else
        Print #lNumArqBAT, "@net use " & mLetraMapeamento & ": /DELETE"
    End If
    Close #lNumArqBAT

    '3 Valida a criação do arquivo
    If gM46V999.gfPreenchido(Dir(lEndArqBAT)) = False Then
        Call gM46V999.gpGraLog(2, "Não foi possivel criar o arquivo P46V001_NETUSE.BAT")
        End
    End If

    Call ExecCmd(gM46V999.gAppPath & "P46V001_NETUSE.BAT", 1)
    If lOpcao = True Then
        If gM46V999.gfPreenchido(Dir(mLetraMapeamento & ":\docstoptransportes\", vbDirectory)) = False Then
            Call gM46V999.gpGraLog(0, "Não foi possivel mapear a unidade " & mLetraMapeamento & ":\." & vbCrLf & vbCrLf & _
                                      "Acesse o servidor de aplicativo (PSLUAPL) e verifique se é possível enxergar a pasta \\SPX20423PSLUYNL\dados\docstoptransportes\." & vbCrLf & "Se estiver tudo ok tente novamente a impressão.")
            Exit Function
        End If
    Else
        Kill gM46V999.gAppPath & "P46V001_NETUSE.BAT"
    End If

    gfGeraBatNetUse = True
    Exit Function
ControleErro:
    Call gM46V999.gpGraLog(2, Err.Description)
    Exit Function
End Function


Public Sub gpInstanciarForm(pobjstcPed As stcA46V702B, pTipoShow As Integer, pTipoDocumento As e_TipDoc)
    Dim lChave As String
    Dim lMensagem As String
    Dim lNomeTemp As String
    Dim lobjWord As Word.Application

    Screen.MousePointer = vbHourglass
    mTipoDocumento = pTipoDocumento


    If mTipoDocumento = e_TipDoc_Cotacao Then
        lNomeTemp = "ESP_" & pobjstcPed.Num_PI & ".doc"
    Else
        lNomeTemp = "ESP_" & pobjstcPed.Num_Apol & "_" & pobjstcPed.Num_PI & ".doc"
    End If
    Set lobjWord = New Word.Application
    Call mfVisualizarEspecificacao(lNomeTemp, pTipoShow, pobjstcPed, lobjWord)

    Screen.MousePointer = vbDefault

End Sub

Private Function mfVisualizarEspecificacao(pNomeArquivo As String, pTypeShow As Integer, pobjstc As stcA46V702B, _
                                           pobjWord As Word.Application) As Boolean

    Dim lobjstcLoc As stcA46V708B    'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
    Dim lobjstcSeguro As stcA46V716B    'Objeto da classe de estrutura stcA46V716B (TAB_PED_LOC_SEGURO).

    Dim NomeTemp As String
    Dim MyPath As String
    Dim lCEP As String
    Dim Aux As String
    Dim i As Integer
    Dim lInd As Integer
    Dim lTextoTmp As String
    Dim lArqTemp As String
    Dim lPesqDigito As String
    Dim clsA99V001 As clsA99V001
    Dim lMensagem As String
    Dim lNovoNome As String
    Dim lUltImpressora As String
    Dim lLocalArqPDF As String
    Dim lDiretorio As String
    Dim objDoc As Word.Document
    Dim objSec As Word.Section
    Dim oShape As Word.Shape
    Dim lCountTable As Long
    Dim lI As Long
    On Error GoTo ControleErro

    'pobjWord.Visible = False
    pobjWord.Application.DisplayAlerts = wdAlertsNone


    NomeTemp = pNomeArquivo

    On Error Resume Next
    Kill NomeTemp
    MkDir gM46V999.gAppPath & "Cotacao\"
    MkDir gM46V999.gAppPath & "Emissao\"
    On Error GoTo ControleErro


    If pobjWord.Documents.Count = 1 Then
        Call mfJuntarArquivos(pobjWord, gM46V999.gPathArquivos & "especificacao2.doc", True, lMensagem, True, _
           , 2.54, 3.7, 3.17, 3.17, 0, 0)

        Call mpReplaceWord(pobjWord.ActiveDocument, "@@APOLICE_RODAPE@@", "Apólice " & Format(pobjstc.Num_Apol, "0000000000") & "-" & pobjstc.mDigitApolice, True)
        Set objDoc = pobjWord.Documents(pobjWord.Documents.Count)
    Else
        pobjWord.Documents.Add
        Set objDoc = pobjWord.Documents(pobjWord.Documents.Count)
        Set objSec = objDoc.Sections(1)
        If Dir(App.Path & "\YasMulti.dot") <> "" Then
            objDoc.CopyStylesFromTemplate Template:=App.Path & "\YasMulti.dot"
        End If
    End If

    If mTipoDocumento = e_TipDoc_Cotacao Then
        On Error Resume Next

        Set oShape = objSec.Headers(wdHeaderFooterEvenPages).Shapes.AddPicture(filename:=App.Path & "\logoYasuda.gif", LinkToFile:=False)     ', Anchor:=objSec.Headers(wdHeaderFooterPrimary).Range)
        With oShape
            .Height = 30
            .Width = 144
            .Left = 0
            .Top = 0
        End With
        Err = 0
        On Error GoTo ControleErro

        'Cabeçalho
        Call gM46V111.gpObtModalidade(pobjstc.Cod_Estip, pobjstc.Cod_Ramo)

        objSec.Headers(wdHeaderFooterPrimary).Range.Font.Name = "Arial"
        objSec.Headers(wdHeaderFooterPrimary).Range.Font.Size = 10
        objSec.Headers(wdHeaderFooterPrimary).Range.Font.bold = True
        For i = 1 To 5
            objSec.Headers(wdHeaderFooterPrimary).Range.InsertAfter vbCrLf
        Next
        objSec.Headers(wdHeaderFooterPrimary).Range.InsertAfter "ESPECIFICAÇÃO"
        objSec.Headers(wdHeaderFooterPrimary).Range.InsertAfter vbCrLf
        objSec.Headers(wdHeaderFooterPrimary).Shapes.AddLine(50, 100, 550, 100).Line.Weight = 1
        objSec.Headers(wdHeaderFooterPrimary).Shapes(1).Line.Weight = 1
        objSec.Headers(wdHeaderFooterPrimary).Range.InsertAfter vbCrLf & pobjstc.gDesc_Ramo
        objSec.Headers(wdHeaderFooterPrimary).Range.InsertAfter vbCrLf
        objSec.Headers(wdHeaderFooterPrimary).Range.InsertAfter vbCrLf & gM46V111.gModalidade


        'Margem
        objDoc.PageSetup.LeftMargin = 50
        objDoc.PageSetup.RightMargin = 50
        objDoc.PageSetup.TopMargin = 70
        objDoc.PageSetup.BottomMargin = 70
        objDoc.PageSetup.HeaderDistance = 25
    End If
    'Cabeçalho
    objDoc.Paragraphs.Add
    RchFinal.TextRTF = ""
    If mTipoDocumento = e_TipDoc_Cotacao Then
        Call mpJuntarTextos("N° COTAÇÃO:", pobjstc.NOSSO_NUMERO, 0, 4)
    Else
        Set clsA99V001 = New clsA99V001
        Set clsA99V001.M46V999 = gM46V999

        Call clsA99V001.DIG_YASUDA(pobjstc.Num_Apol, lPesqDigito, lMensagem)
        Call mpJuntarTextos("", "ESPECIFICAÇÃO ANEXA À APOLICE Nº " & Format(pobjstc.Num_Apol, "0000000000") & "-" & lPesqDigito, 0, 3, True)
        Call mpJuntarTextos("", vbCrLf & vbCrLf, 0, 0)
    End If
    Call mpJuntarTextos("SEGURADO:", pobjstc.Nom_Segurado, 0, 4)

    Call mpSalvarArqTemp(lArqTemp)
    objDoc.Paragraphs(objDoc.Paragraphs.Count).Range.InsertFile lArqTemp


    If pobjstc.ColecaoPedLoc.Count > 0 Then
        Set lobjstcLoc = pobjstc.ColecaoPedLoc.Item(1)
    Else
        Set lobjstcLoc = New stcA46V708B
        Set lobjstcLoc.M46V999 = gM46V999

    End If

    Call mpJuntarTextos("LIMITE MÁXIMO GARANTIA:", Format(pobjstc.Val_IS, "standard"), 0, 2)
    Call mpJuntarTextos("", "", 0, 2)
    Call mpJuntarTextos("VALOR EM RISCO:", Format(lobjstcLoc.Val_Vr, "standard"), 0, 3)

    Call mpSalvarArqTemp(lArqTemp)
    objDoc.Paragraphs(objDoc.Paragraphs.Count).Range.InsertFile lArqTemp

    If pobjstc.ColecaoPedCorr.Count > 0 Then
        Call mpCarregarSubtitulos(objDoc, "CORRETOR")
        Call mpCriarTabelaCorretor(objDoc, pobjstc.ColecaoPedCorr)
    End If

    lCountTable = pobjWord.ActiveDocument.Tables.Count

    If gM46V999.gfPreenchido(pobjstc.objstcEspec.Dsc_Objetos) Then
        Call mpCarregarSubtitulos(objDoc, "OBJETO SEGURADO")
        Call mpJuntarTextoRTF("", pobjstc.objstcEspec.Dsc_Objetos)
        Call mpSalvarArqTemp(lArqTemp)
        objDoc.Paragraphs(objDoc.Paragraphs.Count).Range.InsertFile lArqTemp
        objDoc.Paragraphs(objDoc.Paragraphs.Count).Outdent
    End If

    If gM46V999.gfPreenchido(pobjstc.objstcEspec.Dsc_RCobertos) Then
        Call mpCarregarSubtitulos(objDoc, "RISCOS COBERTOS")
        Call mpJuntarTextoRTF("", pobjstc.objstcEspec.Dsc_RCobertos)
        Call mpSalvarArqTemp(lArqTemp)
        objDoc.Paragraphs(objDoc.Paragraphs.Count).Range.InsertFile lArqTemp
        objDoc.Paragraphs(objDoc.Paragraphs.Count).Outdent
    End If

    'Local de Risco
    Call mpCarregarSubtitulos(objDoc, "LOCAL DE RISCO")
    objDoc.Paragraphs.Add

    lTextoTmp = Trim(lobjstcLoc.Cod_Tip_Logr_Risco) & " "
    lTextoTmp = lTextoTmp & Trim(lobjstcLoc.Nom_Logr_Risco) & " "
    lTextoTmp = lTextoTmp & "nº " & Trim(lobjstcLoc.objstcPedInsp.Num_Logr_Risco) & ", "
    If gM46V999.gfPreenchido(lobjstcLoc.Dsc_Compl_Risco) Then
        lTextoTmp = lTextoTmp & Trim(lobjstcLoc.Dsc_Compl_Risco) & ", "
    End If
    lTextoTmp = lTextoTmp & Trim(lobjstcLoc.Nom_Bairro_Risco) & " - "
    lTextoTmp = lTextoTmp & Trim(lobjstcLoc.Nom_Cid_Risco) & " - "
    lTextoTmp = lTextoTmp & Trim(lobjstcLoc.Sig_UF_Risco) & ", Cep: "
    lCEP = Format(lobjstcLoc.Num_Cep_Risco, "00000000")
    lTextoTmp = lTextoTmp & Left(lCEP, 5) & "-" & Right(lCEP, 3) & " "
    Call mpJuntarTextos("", lTextoTmp, 0, 0)
    Call mpSalvarArqTemp(lArqTemp)
    objDoc.Paragraphs(objDoc.Paragraphs.Count).Range.InsertFile lArqTemp
    objDoc.Paragraphs(objDoc.Paragraphs.Count).Outdent

    'Endereço do segurado .
    Call mpCarregarSubtitulos(objDoc, "ENDEREÇO DO SEGURADO")
    objDoc.Paragraphs.Add
    lTextoTmp = Trim(pobjstc.objstcPedEndereco.Cod_Tip_Logr) & " "
    lTextoTmp = lTextoTmp & Trim(pobjstc.objstcPedEndereco.Nom_Logr) & " "
    lTextoTmp = lTextoTmp & Trim(pobjstc.objstcPedEndereco.Dsc_Compl_Logr) & ", "
    lTextoTmp = lTextoTmp & Trim(pobjstc.objstcPedEndereco.Nom_Bairro) & ", "
    lTextoTmp = lTextoTmp & Trim(pobjstc.objstcPedEndereco.Nom_Cid) & " - "
    lTextoTmp = lTextoTmp & Trim(pobjstc.objstcPedEndereco.Sig_UF) & ", Cep: "
    lCEP = Format(pobjstc.objstcPedEndereco.Num_CEP, "00000000")
    lTextoTmp = lTextoTmp & Left(lCEP, 5) & "-" & Right(lCEP, 3) & " "
    Call mpJuntarTextos("", lTextoTmp, 0, 0)
    Call mpSalvarArqTemp(lArqTemp)
    objDoc.Paragraphs(objDoc.Paragraphs.Count).Range.InsertFile lArqTemp
    objDoc.Paragraphs(objDoc.Paragraphs.Count).Outdent


    If gM46V999.gfPreenchido(pobjstc.objstcEspec.Dsc_Caracteristica) Then
        Call mpCarregarSubtitulos(objDoc, "CARACTERÍSTICAS")
        Call mpJuntarTextoRTF("", pobjstc.objstcEspec.Dsc_Caracteristica)
        Call mpSalvarArqTemp(lArqTemp)
        objDoc.Paragraphs(objDoc.Paragraphs.Count).Range.InsertFile lArqTemp
        objDoc.Paragraphs(objDoc.Paragraphs.Count).Outdent
    End If

    If gM46V999.gfPreenchido(pobjstc.gDesc_RamoAtiv) Then
        Call mpCarregarSubtitulos(objDoc, "OCUPAÇÃO")
        objDoc.Paragraphs.Add
        Call mpJuntarTextos("", pobjstc.gDesc_RamoAtiv, 0, 0)
        Call mpSalvarArqTemp(lArqTemp)
        objDoc.Paragraphs(objDoc.Paragraphs.Count).Range.InsertFile lArqTemp
        objDoc.Paragraphs(objDoc.Paragraphs.Count).Outdent
    End If

    If gM46V999.gfPreenchido(pobjstc.objstcEspec.Dsc_Perimetro) Then
        Call mpCarregarSubtitulos(objDoc, "PERÍMETRO DE COBERTURA")
        objDoc.Paragraphs.Add
        Call mpJuntarTextoRTF("", pobjstc.objstcEspec.Dsc_Perimetro)
        Call mpSalvarArqTemp(lArqTemp)
        objDoc.Paragraphs(objDoc.Paragraphs.Count).Range.InsertFile lArqTemp
        objDoc.Paragraphs(objDoc.Paragraphs.Count).Outdent
    End If

    If gM46V999.gfPreenchido(pobjstc.objstcEspec.Dsc_LocalExposicao) Then
        Call mpCarregarSubtitulos(objDoc, "LOCAL DE EXPOSIÇÃO")
        objDoc.Paragraphs.Add
        Call mpJuntarTextos("", pobjstc.objstcEspec.Dsc_LocalExposicao, 0, 0)
        Call mpSalvarArqTemp(lArqTemp)
        objDoc.Paragraphs(objDoc.Paragraphs.Count).Range.InsertFile lArqTemp
        objDoc.Paragraphs(objDoc.Paragraphs.Count).Outdent
    End If

    If gM46V999.gfPreenchido(pobjstc.objstcEspec.Dsc_Especificacao) Then
        Call mpCarregarSubtitulos(objDoc, "ESPECIFICAÇÃO DO SERVIÇO")
        Call mpJuntarTextoRTF("", pobjstc.objstcEspec.Dsc_Especificacao)
        Call mpSalvarArqTemp(lArqTemp)
        objDoc.Paragraphs(objDoc.Paragraphs.Count).Range.InsertFile lArqTemp
        objDoc.Paragraphs(objDoc.Paragraphs.Count).Outdent
    End If

    If gM46V999.gfPreenchido(pobjstc.objstcEspec.Dsc_Coberturas) Then
        Call mpCarregarSubtitulos(objDoc, "COBERTURAS")
        Call mpJuntarTextoRTF("", pobjstc.objstcEspec.Dsc_Coberturas)
        Call mpSalvarArqTemp(lArqTemp)
        objDoc.Paragraphs(objDoc.Paragraphs.Count).Range.InsertFile lArqTemp
        objDoc.Paragraphs(objDoc.Paragraphs.Count).Outdent
    End If

    If gM46V999.gfPreenchido(pobjstc.objstcEspec.Dsc_Franquia) Then
        Call mpCarregarSubtitulos(objDoc, "FRANQUIA")
        Call mpJuntarTextoRTF("", pobjstc.objstcEspec.Dsc_Franquia)
        Call mpSalvarArqTemp(lArqTemp)
        objDoc.Paragraphs(objDoc.Paragraphs.Count).Range.InsertFile lArqTemp
        objDoc.Paragraphs(objDoc.Paragraphs.Count).Outdent
    End If

    Call mpJuntarTextos("PRAZO DO SEGURO:", gM46V999.gfForData(2, pobjstc.Dat_Ini_Vig) & "   à   " & gM46V999.gfForData(2, pobjstc.Dat_Ter_Vig), 0, 2)

    If pobjstc.objstcEspec.Dat_Obra <> 0 Then
        Call mpJuntarTextos("PRAZO DA OBRA:", gM46V999.gfForData(2, pobjstc.objstcEspec.Dat_Obra), 0, 2)
    End If
    If pobjstc.objstcEspec.Dat_Manut <> 0 Then
        Call mpJuntarTextos("PRAZO DA MANUTENÇÃO:", gM46V999.gfForData(2, pobjstc.objstcEspec.Dat_Manut), 0, 1)
    End If
    'If lobjstcLoc.objstcPedClau.Cod_Clau <> 0 Then
    '    Aux = Val(Left(lobjstcLoc.objstcPedClau.Descricao, 14))
    '    If Len(Aux) <= 11 Then
    '        Call gM46V999.gpForCPF(1, Aux, Aux)
    '    Else
    '        Call gM46V999.gpForCNPJ(1, Aux, Aux)
    '    End If
    '    Call mpJuntarTextos("BENEFICIÁRIO:", Trim(Mid(lobjstcLoc.objstcPedClau.Descricao, 15, 60)) & " - CPF nº" & Aux, 0, 3)
    'End If
    If pobjstc.ColecaoPedCoss.Count > 0 Then
        RchFinal.SelText = vbCrLf
        Call mpJuntarTextos("COSSEGURO:", "", 0, 0)
        Call mpCriarTabelaCosseguro(objDoc, pobjstc.ColecaoPedCoss)
    End If

    Call mpSalvarArqTemp(lArqTemp)
    objDoc.Paragraphs(objDoc.Paragraphs.Count).Range.InsertFile lArqTemp

    If gM46V999.gfPreenchido(pobjstc.objstcEspec.Nom_Contratante) Then
        Call mpCarregarSubtitulos(objDoc, "CONTRATANTE")
        objDoc.Paragraphs.Add
        Call mpJuntarTextos("", pobjstc.objstcEspec.Nom_Contratante, 0, 0)
        Call mpSalvarArqTemp(lArqTemp)
        objDoc.Paragraphs(objDoc.Paragraphs.Count).Range.InsertFile lArqTemp
        objDoc.Paragraphs(objDoc.Paragraphs.Count).Outdent
    End If

    If gM46V999.gfPreenchido(pobjstc.objstcEspec.Dsc_Clausulas) Then
        Call mpCarregarSubtitulos(objDoc, "CLÁUSULAS")
        Call mpJuntarTextoRTF("", pobjstc.objstcEspec.Dsc_Clausulas)
        Call mpSalvarArqTemp(lArqTemp)
        objDoc.Paragraphs(objDoc.Paragraphs.Count).Range.InsertFile lArqTemp
        objDoc.Paragraphs(objDoc.Paragraphs.Count).Outdent
    End If

    If gM46V999.gfPreenchido(pobjstc.objstcEspec.Dsc_Condicao) Then
        Call mpCarregarSubtitulos(objDoc, "CONDIÇÕES ESPECIAIS")
        Call mpJuntarTextoRTF("", pobjstc.objstcEspec.Dsc_Condicao)
        Call mpSalvarArqTemp(lArqTemp)
        objDoc.Paragraphs(objDoc.Paragraphs.Count).Range.InsertFile lArqTemp
        objDoc.Paragraphs(objDoc.Paragraphs.Count).Outdent
    End If

    If lobjstcLoc.ColecaoPedSeguro.Count > 0 Then
        Call mpCarregarSubtitulos(objDoc, "OUTROS SEGUROS")
        objDoc.Paragraphs.Add
        For Each lobjstcSeguro In lobjstcLoc.ColecaoPedSeguro
            Call mpJuntarTextos("", lobjstcSeguro.Dsc_Outro_Seguro & " - Apólice: " & lobjstcSeguro.Num_Apol_Ant & " - Item: " & lobjstcSeguro.Num_Item_Ant, 0, 0)
        Next lobjstcSeguro
        Call mpSalvarArqTemp(lArqTemp)
        objDoc.Paragraphs(objDoc.Paragraphs.Count).Range.InsertFile lArqTemp
        objDoc.Paragraphs(objDoc.Paragraphs.Count).Outdent
    End If

    If gM46V999.gfPreenchido(pobjstc.objstcEspec.Dsc_OutrasInf) Then
        Call mpCarregarSubtitulos(objDoc, "OUTRAS INFORMAÇÕES")
        Call mpJuntarTextoRTF("", pobjstc.objstcEspec.Dsc_OutrasInf)
        Call mpSalvarArqTemp(lArqTemp)
        objDoc.Paragraphs(objDoc.Paragraphs.Count).Range.InsertFile lArqTemp
        objDoc.Paragraphs(objDoc.Paragraphs.Count).Outdent
    End If

    If gM46V999.gfPreenchido(pobjstc.objstcEspec.Dsc_Observacao) Then
        Call mpCarregarSubtitulos(objDoc, "OBSERVAÇÃO")
        Call mpJuntarTextoRTF("", pobjstc.objstcEspec.Dsc_Observacao)
        Call mpSalvarArqTemp(lArqTemp)
        objDoc.Paragraphs(objDoc.Paragraphs.Count).Range.InsertFile lArqTemp
        objDoc.Paragraphs(objDoc.Paragraphs.Count).Outdent
    End If

    For lI = lCountTable + 1 To pobjWord.ActiveDocument.Tables.Count
        Call pobjWord.ActiveDocument.Tables.Item(lI).Rows.SetLeftIndent(CentimetersToPoints(0), wdAdjustNone)
    Next lI
    Call mpCarregarSubtitulos(objDoc, "PRÊMIO LÍQUIDO")
    objDoc.Paragraphs.Add
    Call mpJuntarTextos("", Format(pobjstc.VAL_PR, "standard"), 0, 5)
    Call mpSalvarArqTemp(lArqTemp)
    objDoc.Paragraphs(objDoc.Paragraphs.Count).Range.InsertFile lArqTemp
    objDoc.Paragraphs(objDoc.Paragraphs.Count).Outdent

    If mTipoDocumento = e_TipDoc_Cotacao Then
        objDoc.Paragraphs.RightIndent = 0

        objDoc.PageSetup.FooterDistance = 72

        'Rodapé
        objSec.Footers(wdHeaderFooterPrimary).Shapes.AddLine 50, 710, 550, 710
        objSec.Footers(wdHeaderFooterPrimary).Range.Font.Name = "Arial"
        objSec.Footers(wdHeaderFooterPrimary).Range.Font.Size = 8
        objSec.Footers(wdHeaderFooterPrimary).Range.Text = "Apolice " & Format(pobjstc.Num_Apol, "0000000000") & Space(10) & "Endosso " & Format(pobjstc.Num_Endo, "000000")


        With objSec.Footers(wdHeaderFooterPrimary).PageNumbers
            .Add PageNumberAlignment:=wdAlignPageNumberCenter
            .StartingNumber = 1
            .HeadingLevelForChapter = 0
            .NumberStyle = wdPageNumberStyleLowercaseRoman
            .IncludeChapterNumber = True
            .HeadingLevelForChapter = 0
            .ChapterPageSeparator = wdSeparatorEnDash
        End With

        pobjWord.Application.Caption = "Especificação"

    End If


    Set objSec = objDoc.Sections(1)
    pobjWord.ActiveWindow.View.Type = wdPageView
    pobjWord.ActiveWindow.View.SeekView = wdSeekPrimaryFooter
    pobjWord.ActiveDocument.Sections(1).Footers(1).Range.Fields.Update
    If pobjWord.ActiveDocument.Sections.Count > 1 Then
        pobjWord.ActiveDocument.Sections(2).Footers(1).Range.Fields.Update
    End If
    If pobjWord.ActiveDocument.Sections.Count > 2 Then
        pobjWord.ActiveDocument.Sections(3).Footers(1).Range.Fields.Update
    End If


    objDoc.SaveAs NomeTemp
    pobjWord.Documents.Save
    'pobjWord.ActiveWindow.View.Zoom.PageFit = wdPageFitBestFit
    objDoc.Protect (wdAllowOnlyComments)
    Select Case pTypeShow
    Case 1, 3    'Visualizar , salva e gera o PDF
        If pTypeShow <> 3 Then
            pobjWord.Visible = True
            pobjWord.Application.Visible = True
            pobjWord.Application.DisplayStatusBar = True
            pobjWord.Application.WindowState = wdWindowStateMaximize
            pobjWord.Application.CommandBars("standard").Visible = True
        End If
        objDoc.Save
        If mTipoDocumento = e_TipDoc_Cotacao Then
            lDiretorio = App.Path & "\Cotacao\"
            lNovoNome = Replace(NomeTemp, gM46V999.gAppPath, lDiretorio)
        Else
            lDiretorio = App.Path & "\Emissao\"
            lNovoNome = Replace(NomeTemp, gM46V999.gAppPath, lDiretorio)
        End If
        On Error Resume Next
        Err.Clear
        Kill lNovoNome
        If Err.Number <> 53 And Err.Number > 0 Then
            Call gM46V999.gpGraLog(1, "Arquivo em uso (" & lNovoNome & ") . Por favor fechar o arquivo e gerar novamente.")
            Exit Function
        End If
        On Error GoTo ControleErro
        objDoc.SaveAs lNovoNome
        On Error Resume Next
        If pTypeShow = 3 Then
            If Not mfImprimir(pobjWord, pobjstc, lMensagem, False) Then GoTo ControleErro
            pobjWord.Documents.Close
            pobjWord.Quit
        End If
        On Error GoTo ControleErro
    Case 2  'Imprimir
        pobjWord.PrintOut
    End Select

    Set pobjWord = Nothing
    Set objDoc = Nothing
    mfVisualizarEspecificacao = True
    Exit Function
ControleErro:
    If Err <> 429 Then
        Call gM46V999.gpGraLog(1, Str$(Err) & " " & Error$)
    Else
        Resume Next
    End If
End Function


Function gfForNormalDot(pobjWord As Word.Application, _
                        ByRef pMsgErr, _
                        Optional pRestaurar As Boolean = False) As Boolean
'################################################################################
'#      Objetivo:
'#               Formatar o documento de modelo do MS Word YasMulti.dot
'#               para o padrão. Prevenindo problemas de resultados não previstos
'#               Ex.: Impressão com todas as linhas sublinhadas.
'#      Variáveis:
'#                pobjWord .....Objeto utilizado na impressão (enviar já iniciado...)
'#                pMsgErr ......Mensagem de retorno de erro (caso ocorra.)
'#                pRestaurar ...=FALSE "Formata p/ modelo padrão" (rodar antes de iniciar a criação do documento)
'#                             .=TRUE  "Restaura as configurações do usuário, case se aplique."
'#                             (rodar após a criação do documento)
'#
'#                             OBS.: Não finalizar o objeto "pobjWord" antes de restaurar o formato do usuário
'#                             (antes do .QUIT ou descarregar a variável)
'################################################################################
    Dim lObjNormalDot As Object
    Dim lobjDocDummy As Word.Document
    Dim lobjSecDummy As Word.Section
    Static lsUserFontUnderline As WdUnderline
    Static lsMudouModelo As Boolean
    'Dim lSalvarNormalDot            As Boolean
    Dim mNomeModelo As String

    Err = 0
    On Error GoTo Func_Erro
    gfForNormalDot = False
    'lSalvarNormalDot = False

    mNomeModelo = App.Path & "\YasMulti.dot"
    If Dir(mNomeModelo) = "" Then
        MsgBox "Modelo padrão não encontrado! Utilizando modelo do word."
        mNomeModelo = "Normal.dot"
    End If

    pobjWord.Application.DisplayAlerts = wdAlertsNone
    If pobjWord.Documents.Count = 1 Then
        pobjWord.Documents.Add
        Set lobjDocDummy = pobjWord.Documents(1)    '***O documento recém adicionado é sempre o 1º
        lobjDocDummy.Sections.Add
        Set lobjSecDummy = lobjDocDummy.Sections(2)
    Else
        pobjWord.Documents.Add
        Set lobjDocDummy = pobjWord.Documents(1)    '***O documento recém adicionado é sempre o 1º
        Set lobjSecDummy = lobjDocDummy.Sections(1)
    End If

    lobjDocDummy.CopyStylesFromTemplate Template:=mNomeModelo
    lobjDocDummy.AttachedTemplate = mNomeModelo

    pobjWord.ActiveDocument.UpdateStyles
    lobjDocDummy.Close SaveChanges:=wdDoNotSaveChanges

    'Formata estilo "Normal" do documento global NORMAL.DOT
    'If pobjWord.Documents.Count >= 1 Then
    'Set lObjNormalDot = pobjWord.ActiveDocument.AttachedTemplate.OpenAsDocument
    'With lObjNormalDot.Styles(wdStyleNormal).Font
    '    If pRestaurar = False Then
    '        '--Sublinhado
    '        If .Underline <> wdUnderlineNone Then
    '            'Guarda o tipo definido pelo usuário (utilizado posteriormente p/ restaurar)
    '            lsUserFontUnderline = .Underline
    '            lsMudouModelo = True
    '            'Ajusta p/ o formato padrão
    '            .Underline = wdUnderlineNone
    '            lSalvarNormalDot = True
    '        End If
    '    Else
    '        If lsMudouModelo = True Then
    '            '--Sublinhado
    '            If .Underline <> lsUserFontUnderline Then
    '                'Ajusta p/ o formato do usuário
    '                .Underline = lsUserFontUnderline
    '                lSalvarNormalDot = True
    '            End If
    '            'Reinicia o flag
    '            lsMudouModelo = False
    '        End If
    '    End If
    'End With
    'If lSalvarNormalDot = True Then
    '    lObjNormalDot.Close SaveChanges:=wdSaveChanges
    '    pobjWord.ActiveDocument.UpdateStyles
    'Else
    '    lObjNormalDot.Close SaveChanges:=wdDoNotSaveChanges
    'End If
    'lobjDocDummy.Close SaveChanges:=wdDoNotSaveChanges
    'End If

    Set lobjDocDummy = Nothing
    Set lobjSecDummy = Nothing
    gfForNormalDot = True
Func_Saida:
    Err = 0
    On Error GoTo 0
    Exit Function

Func_Erro:
    pMsgErr = "gfForNormalDot Erro " & Err & " - " & Error
    On Error Resume Next
    Set lobjDocDummy = Nothing
    Set lobjSecDummy = Nothing
    On Error GoTo 0
    GoTo Func_Saida
End Function

Private Sub mpJuntarFormatos(ByVal NomeLabel, ByVal ValorRich As RichTextBox)

    If Trim(ValorRich.Text) = "" Then Exit Sub

    RchFinal.SelStart = Len(RchFinal.Text)
    RchFinal.SelBold = True
    RchFinal.SelFontSize = 10
    RchFinal.SelColor = &H0&
    RchFinal.SelFontName = gM46V111.gFonteEspecificacao
    RchFinal.SelText = vbCrLf
    RchFinal.SelText = NomeLabel
    RchFinal.SelText = vbCrLf
    RchFinal.SelBold = False
    RchFinal.SelRTF = ValorRich.TextRTF

End Sub

Private Sub mpJuntarTextoRTF(ByVal NomeLabel, ByVal ValorRich As String)
    Dim CampoRTF As RichTextBox
    Set CampoRTF = Me.RchTemp

    CampoRTF.TextRTF = ValorRich

    If Trim(CampoRTF.Text) = "" Then Exit Sub

    RchFinal.SelStart = Len(RchFinal.Text)
    RchFinal.SelBold = True
    RchFinal.SelFontSize = 10
    RchFinal.SelColor = &H0&
    RchFinal.SelFontName = gM46V111.gFonteEspecificacao
    RchFinal.SelText = vbCrLf
    If NomeLabel <> "" Then
        RchFinal.SelText = NomeLabel
        RchFinal.SelText = vbCrLf
    End If
    RchFinal.SelBold = False
    RchFinal.SelRTF = CampoRTF.TextRTF

End Sub

Private Sub mpJuntarTextos(ByVal NomeLabel, ByVal Texto As String, ByVal Quebra As Integer, ByVal Tabs As Integer, Optional pBold As Boolean = False)
    Dim i As Integer
    RchFinal.SelStart = Len(RchFinal.Text)
    RchFinal.SelBold = True
    RchFinal.SelFontSize = 10
    RchFinal.SelColor = &H0&
    RchFinal.SelFontName = gM46V111.gFonteEspecificacao
    RchFinal.SelText = NomeLabel
    RchFinal.SelBold = pBold
    If Quebra > 0 Then
        For i = 1 To Quebra
            RchFinal.SelText = vbCrLf
        Next
    End If
    If Tabs > 0 Then
        For i = 1 To Tabs
            RchFinal.SelText = vbTab
        Next
    End If
    RchFinal.SelFontSize = 9
    RchFinal.SelText = Texto
    RchFinal.SelText = vbCrLf
End Sub

Private Sub mpDeletarTemp(ByVal Diretorio As String)
    On Error Resume Next
    Dim NomeArq As String

    NomeArq = Dir(Diretorio & "#Tmp*.ytf", vbArchive)
    Do While NomeArq <> ""   ' Inicia o loop.
        If NomeArq <> "." And NomeArq <> ".." Then
            Kill Diretorio & NomeArq
        End If
        NomeArq = Dir   ' Obtém a próxima entrada.
    Loop
    Err = 0
End Sub

Private Sub mpCarregarSubtitulos(ByRef pWord As Word.Document, pSubtitulo As String)
    pWord.Paragraphs.Add
    With pWord.Paragraphs(pWord.Paragraphs.Count - 1).Range
        .InsertBefore pSubtitulo
        .bold = True
        .ParagraphFormat.Alignment = wdAlignParagraphCenter


        '        .ParagraphFormat.Shading.Texture = wdTexture25Percent
        '        .ParagraphFormat.Shading.ForegroundPatternColorIndex = wdAuto
        '        .ParagraphFormat.Shading.BackgroundPatternColorIndex = wdWhite

        '        .ParagraphFormat.Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        '        .ParagraphFormat.Borders(wdBorderRight).LineStyle = wdLineStyleNone
        '        .ParagraphFormat.Borders(wdBorderTop).LineStyle = wdLineStyleNone
        '        .ParagraphFormat.Borders(wdBorderBottom).LineStyle = wdLineStyleNone

        .ParagraphFormat.Borders(wdBorderTop).LineStyle = wdLineStyleSingle
        .ParagraphFormat.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle

        .Font.Name = gM46V111.gFonteEspecificacao
        .Font.Size = 10

    End With

End Sub

Private Sub mpSalvarArqTemp(ByRef pNom_Arquivo As String)
    Dim MyPath As String
    On Error Resume Next
    MyPath = "c:\temp\"
    MkDir MyPath
    On Error GoTo objError

    pNom_Arquivo = "#Tmp" & Format(Now(), "hh:mm:ss") & ".ytf"
    pNom_Arquivo = MyPath & Replace(pNom_Arquivo, ":", "")

    'Salvar arquivo Montado
    RchFinal.SaveFile pNom_Arquivo
    RchFinal.TextRTF = ""
    Exit Sub
objError:
    Call gM46V999.gpGraLog(1, Err.Description)
End Sub


Private Sub mpCriarTabelaCorretor(pWord As Word.Document, pColecaoCorretor As Collection)
    Dim lobjstcCorretor As stcA46V705B    'Objeto da classe de estrutura stcA46V705B (TAB_PED_CORR)
    Dim lCount As Long

    pWord.Paragraphs.Add
    pWord.Paragraphs(pWord.Paragraphs.Count).Range.Font.Name = gM46V111.gFonteEspecificacao

    pWord.Tables.Add pWord.Range(pWord.Range.End - 1), NumRows:=pColecaoCorretor.Count + 1, NumColumns:=3
    pWord.Tables(pWord.Tables.Count).Columns(1).Cells(1).Range.Text = "Corretor"
    pWord.Tables(pWord.Tables.Count).Columns(2).Cells(1).Range.Text = "Código"
    pWord.Tables(pWord.Tables.Count).Columns(3).Cells(1).Range.Text = "Comissão"
    lCount = 1
    For Each lobjstcCorretor In pColecaoCorretor
        lCount = lCount + 1
        pWord.Tables(pWord.Tables.Count).Columns(1).Cells(lCount).Range.Text = lobjstcCorretor.gNom_Corretor
        pWord.Tables(pWord.Tables.Count).Columns(2).Cells(lCount).Range.Text = lobjstcCorretor.Cod_Corr
        pWord.Tables(pWord.Tables.Count).Columns(3).Cells(lCount).Range.Text = Format(lobjstcCorretor.Per_Corr, "standard") & "%"
    Next lobjstcCorretor
    pWord.Paragraphs.Add

End Sub
Private Sub mpCriarTabelaCosseguro(pWord As Word.Document, pColecaoCosseguro As Collection)
    Dim lobjstcCosseguro As stcA46V706B
    Dim lCount As Long

    pWord.Tables.Add pWord.Range(pWord.Range.End - 1), NumRows:=pColecaoCosseguro.Count + 1, NumColumns:=3

    pWord.Tables(pWord.Tables.Count).Columns(1).Cells(1).Range.Text = "Companhia seguradora"
    pWord.Tables(pWord.Tables.Count).Columns(2).Cells(1).Range.Text = "Código"
    pWord.Tables(pWord.Tables.Count).Columns(3).Cells(1).Range.Text = "Participação"
    lCount = 1
    For Each lobjstcCosseguro In pColecaoCosseguro
        lCount = lCount + 1
        pWord.Tables(pWord.Tables.Count).Columns(1).Cells(lCount).Range.Text = lobjstcCosseguro.Nom_Seguradora
        pWord.Tables(pWord.Tables.Count).Columns(2).Cells(lCount).Range.Text = lobjstcCosseguro.Cod_Coss
        pWord.Tables(pWord.Tables.Count).Columns(3).Cells(lCount).Range.Text = Format(lobjstcCosseguro.Per_Coss, "standard") & "%"
    Next lobjstcCosseguro
    pWord.Paragraphs(pWord.Paragraphs.Count - 1).Range.Font.Name = gM46V111.gFonteEspecificacao

End Sub

Private Sub mpReplaceWord(ByRef pWord As Word.Document, pLocalizar As String, _
                          pSubstituir As String, Optional pLocalizarHeader As Boolean = False)
    Dim lLocalizou As Boolean
    Dim lI As Long

    pWord.Range.Find.ClearFormatting
    pWord.Range.Find.Replacement.ClearFormatting
    lLocalizou = pWord.Range.Find.Execute(pLocalizar, False, False, False, _
                                          False, False, True, wdFindContinue, False, pSubstituir, wdReplaceAll)

    If pLocalizarHeader = True Then
        pWord.ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
        For lI = 1 To pWord.Sections(1).Footers(wdHeaderFooterEvenPages).Shapes.Count
            If UCase(pWord.Sections(1).Footers(wdHeaderFooterEvenPages).Shapes(lI).Name) Like "*TEXT*" Then
                Call pWord.Sections(1).Footers(wdHeaderFooterEvenPages).Shapes(lI).Select
                If Replace(pWord.Sections(1).Footers(wdHeaderFooterEvenPages).Shapes(lI).TextFrame.TextRange, vbCr, "") = pLocalizar Then
                    pWord.Sections(1).Footers(wdHeaderFooterEvenPages).Shapes(lI).TextFrame.TextRange = pSubstituir
                    Exit For
                End If
            End If
        Next lI
        If pWord.Sections.Count > 1 Then
            For lI = 1 To pWord.Sections(2).Footers(wdHeaderFooterEvenPages).Shapes.Count
                If UCase(pWord.Sections(2).Footers(wdHeaderFooterEvenPages).Shapes(lI).Name) Like "*TEXT*" Then
                    Call pWord.Sections(2).Footers(wdHeaderFooterEvenPages).Shapes(lI).Select
                    If Replace(pWord.Sections(2).Footers(wdHeaderFooterEvenPages).Shapes(lI).TextFrame.TextRange, vbCr, "") = pLocalizar Then
                        pWord.Sections(2).Footers(wdHeaderFooterEvenPages).Shapes(lI).TextFrame.TextRange = pSubstituir
                        Exit For
                    End If
                End If
            Next lI
        End If
        If pWord.Sections.Count > 2 Then
            For lI = 1 To pWord.Sections(3).Footers(wdHeaderFooterEvenPages).Shapes.Count
                If UCase(pWord.Sections(3).Footers(wdHeaderFooterEvenPages).Shapes(lI).Name) Like "*TEXT*" Then
                    Call pWord.Sections(3).Footers(wdHeaderFooterEvenPages).Shapes(lI).Select
                    If Replace(pWord.Sections(3).Footers(wdHeaderFooterEvenPages).Shapes(lI).TextFrame.TextRange, vbCr, "") = pLocalizar Then
                        pWord.Sections(3).Footers(wdHeaderFooterEvenPages).Shapes(lI).TextFrame.TextRange = pSubstituir
                        Exit For
                    End If
                End If
            Next lI
        End If
    End If


End Sub

Private Sub mpEmitirDocumentos()
    Dim lI As Long
    Dim lUltData As Double
    Dim lUltUsuarioa As String
    Dim lCount As Long
    Dim lArqApolice As String
    Dim lDiretorio As String
    Dim lArqEspecif As String
    Dim lHoraInicial As String
    Dim lHoraFinal As String
    Dim lObjstcPed As stcA46V702B
    Dim lobjPedEspec As clsA46V721A
    Dim lobjstcPedEspec As stcA46V721B
    Dim lobjWord As Word.Application


    lCount = 0
    lHoraInicial = Now
    For lI = 0 To lstLista.ListCount - 1
        If lstLista.Selected(lI) = True Then
            lCount = lCount + 1
            Set lObjstcPed = mobjPed.Colecao.Item(Format(lstLista.ItemData(lI), "0000000000"))
            mTipoDocumento = e_TipDoc_Proposta
            lArqApolice = gM46V999.gAppPath & "Apol_" & lObjstcPed.objstcCtrlEmis.Num_Apol & "_" & lObjstcPed.Num_PI & ".doc"
            lArqEspecif = gM46V999.gAppPath & "ESP_" & lObjstcPed.objstcCtrlEmis.Num_Apol & "_" & lObjstcPed.Num_PI & ".doc"
            Call mfSeparador(lObjstcPed, 1)
            Set lobjWord = New Word.Application
            If lObjstcPed.objstcCtrlEmis.Num_Endo > 0 Then
                panAviso.Caption = "Gerando arquivos do endosso " & Format(lObjstcPed.objstcCtrlEmis.Num_Apol, "0000000000") & "/" & Format(lObjstcPed.objstcCtrlEmis.Num_Endo, "000000") & " ( PI : " & lObjstcPed.Num_PI & ")"
                'Gera arquivo de endosso
                If Not mfGerarArqEndosso(lArqApolice, lObjstcPed, lobjWord) Then Exit Sub
            Else
                panAviso.Caption = "Gerando arquivos da apólices " & Format(lObjstcPed.objstcCtrlEmis.Num_Apol, "0000000000") & " ( PI : " & lObjstcPed.Num_PI & ")"
                'Gera arquivo de apólice
                If Not mfGerarArqApolice(lArqApolice, lObjstcPed, lobjWord) Then Exit Sub
            End If

            'Incluo os dados de cosseguro, beneficiário, fracionamento, outros seguros, bens
            If Not mfGerarDemaisDocs(lObjstcPed, lobjWord) Then Exit Sub
            'Gera arquivo de especificação
            If Not mfVisualizarEspecificacao(lArqApolice, 3, lObjstcPed, lobjWord) Then Exit Sub
            'Somente para o segurado (1) (finanças quando estrangeira - (3)) será impresso as cláusulas
            If Not mfVisualizarClausulas(lArqApolice, lArqEspecif, lObjstcPed) Then Exit Sub

            'Gera a via do corretor
            Call mfSeparador(lObjstcPed, 2)
            If Not mfGerarFichaCorretor(lObjstcPed) Then Exit Sub

            'Excluir arquivos
            On Error Resume Next
            Kill lArqApolice
            Kill lArqEspecif
            On Error GoTo 0
            panAviso.Caption = "Processados " & lI + 1 & " de " & mobjPed.Colecao.Count


            Set lobjPedEspec = New clsA46V721A
            Set lobjPedEspec.M46V999 = gM46V999

            lObjstcPed.objstcEspec.Ind_Impressao = 1

            'If Not lobjPedEspec.mfExcluir(lobjPedEspec) Then
            '    Call gM46V999.gpGraLog(1, lobjPedEspec.UltimoErro)
            '    Exit Sub
            'End If
            'If Not lobjPedEspec.mfIncluir(lObjstcPed) Then
            '    Call gM46V999.gpGraLog(1, lobjPedEspec.UltimoErro)
            '    Exit Sub
            'End If



        End If
    Next lI

    lHoraFinal = Now
    Call gM46V999.gpGraLog(1, "Impressão formulários Diversos." & vbCrLf & _
                              "Total de documentos impressos: " & lCount & _
                              vbCrLf & vbCrLf & _
                              "Início: " & lHoraInicial & vbCrLf & _
                              "Término: " & lHoraFinal)
    Unload Me


End Sub
Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub mpCarregarCelulaTabela(pobjWord As Word.Application, pLinha As Long, _
                                   lColuna As Long, pTexto As String, pBold As Boolean, _
                                   pAlinhamento As WdParagraphAlignment, Optional pTamanhoColuna As Double)

    With pobjWord.ActiveDocument
        .Tables(.Tables.Count).Columns(lColuna).Cells(pLinha).Range.Text = pTexto
        If Not IsMissing(pTamanhoColuna) Then
            .Tables(.Tables.Count).Columns(lColuna).SetWidth pobjWord.CentimetersToPoints(pTamanhoColuna), wdAdjustNone
        End If
        .Tables(.Tables.Count).Columns(lColuna).Cells(pLinha).Range.bold = pBold

        .Tables(.Tables.Count).Columns(lColuna).Cells(pLinha).Range.ParagraphFormat.Alignment = pAlinhamento
    End With

End Sub

Private Sub mpCriarTabela(pobjWord As Word.Application, pQtdeLinhas As Long, _
                          pQtdeColunas As Long, pRetirarBorda As Boolean)

    With pobjWord.ActiveDocument
        .Tables.Add .Range(.Range.End - 1), NumRows:=pQtdeLinhas, NumColumns:=pQtdeColunas
        .Tables(.Tables.Count).Rows.SetLeftIndent pobjWord.CentimetersToPoints(0), wdAdjustNone
        If pRetirarBorda = True Then
            .Tables(.Tables.Count).Borders(wdBorderBottom).LineStyle = wdLineStyleNone
            .Tables(.Tables.Count).Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
            .Tables(.Tables.Count).Borders(wdBorderLeft).LineStyle = wdLineStyleNone
            .Tables(.Tables.Count).Borders(wdBorderRight).LineStyle = wdLineStyleNone
            .Tables(.Tables.Count).Borders(wdBorderTop).LineStyle = wdLineStyleNone
            .Tables(.Tables.Count).Borders(wdBorderVertical).LineStyle = wdLineStyleNone
        End If
    End With
End Sub
Private Sub mpNovaPaginaOutrosDocs(pobjWord As Word.Application, pTitulo As String, Optional pNovaPag As Boolean = True)

    With pobjWord.ActiveDocument
        If pNovaPag = True Then
            .Range(.Range.End - 1).Select
            .Paragraphs.Add
            .Range(.Range.End - 1).InsertBreak Type:=wdPageBreak
        End If
        .Range(.Range.End - 1).Select
        .Paragraphs.Add
        If gM46V999.gfPreenchido(pTitulo) Then
            .Paragraphs(.Paragraphs.Count).Range.Font.bold = True
            .Paragraphs(.Paragraphs.Count).Range.Font.Name = "ARIAL"
            .Paragraphs(.Paragraphs.Count).Range.Font.Size = 10
            .Paragraphs(.Paragraphs.Count).Range.Text = pTitulo
            .Paragraphs.Add
            .Paragraphs.Add
            Call mpCriarTabela(pobjWord, 1, 1, True)
            .Tables(.Tables.Count).Borders(wdBorderTop).LineStyle = wdLineStyleSingle
            .Tables(.Tables.Count).Borders(wdBorderTop).LineWidth = wdLineWidth150pt
            .Paragraphs.Add
        End If
        .Paragraphs.Add
        .Paragraphs(.Paragraphs.Count).Range.Font.bold = False
        .Paragraphs(.Paragraphs.Count).Range.Font.Size = 9
        .Range(.Range.End - 1).Select
    End With
    mContadorPagina = mContadorPagina + 1
End Sub

Private Function mfGerarDemaisDocs(pobjstcPed As stcA46V702B, pobjWord As Word.Application) As Boolean
    Dim lobjstcLoc As stcA46V708B
    Dim lObjstcBens As stcA46V715B
    Dim lobjstcSeguro As stcA46V716B
    Dim lobjstcCoss As stcA46V706B
    Dim lobjstcBenef As stcA46V726B
    Dim lObjstcVenc As stcA46V713B

    Dim lCount As Long
    Dim lPercentual As Double
    Dim lValTotalBem As Double
    Dim lAuxiliar As String
    Dim lValorParc As Variant
    Dim lMensagem As String
    Dim lNomeArqFrac As String
    Dim lArquivo As Integer
    Dim lLinhas As Long
    Dim lRegistro As String
    Dim lNovaPag As Boolean

    mContadorPagina = 0
    pobjWord.Visible = True
    Call mfJuntarArquivos(pobjWord, gM46V999.gPathArquivos & "especificacao.doc", True, lMensagem, True, _
       , 2.54, 3.7, 3.17, 3.17, 0, 0)

    Call mpReplaceWord(pobjWord.ActiveDocument, "@@APOLICE_RODAPE@@", "Apólice " & Format(pobjstcPed.Num_Apol, "0000000000") & "-" & pobjstcPed.mDigitApolice, True)
    lNovaPag = False
    With pobjWord.ActiveDocument
        'Distribuição de cosseguro
        If pobjstcPed.ColecaoPedCoss.Count > 0 Then
            Call mpNovaPaginaOutrosDocs(pobjWord, "DISTRIBUIÇÃO DE COSSEGURO", lNovaPag)
            lNovaPag = True
            Call mpCriarTabela(pobjWord, pobjstcPed.ColecaoPedCoss.Count + 3, 7, True)

            Call mpCarregarCelulaTabela(pobjWord, 1, 1, "Companhia(s)", True, 0, 4.77)
            Call mpCarregarCelulaTabela(pobjWord, 1, 2, "Num. Ordem", True, 0, 2.25)
            Call mpCarregarCelulaTabela(pobjWord, 1, 3, "%Part", True, 0, 1.25)
            Call mpCarregarCelulaTabela(pobjWord, 1, 4, "Garantia", True, 0, 1.6)
            Call mpCarregarCelulaTabela(pobjWord, 1, 5, "Importância Segurada", True, 0, 3.6)
            Call mpCarregarCelulaTabela(pobjWord, 1, 6, "Prêmio Líquido", True, 0, 2.6)
            Call mpCarregarCelulaTabela(pobjWord, 1, 7, "Adic. Fracionamento", True, 0, 3.4)

            lCount = 1
            For Each lobjstcCoss In pobjstcPed.ColecaoPedCoss
                lCount = lCount + 1
                Call mpCarregarCelulaTabela(pobjWord, lCount, 1, lobjstcCoss.Cod_Coss & " - " & lobjstcCoss.gNom_Seguradora, False, wdAlignParagraphLeft)
                Call mpCarregarCelulaTabela(pobjWord, lCount, 2, lobjstcCoss.gNumOrdem, False, wdAlignParagraphLeft)
                Call mpCarregarCelulaTabela(pobjWord, lCount, 3, Format(lobjstcCoss.Per_Coss, "standard"), False, wdAlignParagraphRight)
                Call mpCarregarCelulaTabela(pobjWord, lCount, 5, "0,00", False, wdAlignParagraphRight)
                Call mpCarregarCelulaTabela(pobjWord, lCount, 6, "0,00", False, wdAlignParagraphRight)
                Call mpCarregarCelulaTabela(pobjWord, lCount, 7, "0,00", False, wdAlignParagraphRight)
                lPercentual = lPercentual + lobjstcCoss.Per_Coss
            Next lobjstcCoss
            lCount = lCount + 1
            Call mpCarregarCelulaTabela(pobjWord, lCount, 1, "641 - YASUDA SEGUROS", False, wdAlignParagraphLeft)
            Call mpCarregarCelulaTabela(pobjWord, lCount, 3, Format(100 - lPercentual, "standard"), False, wdAlignParagraphRight)
            Call mpCarregarCelulaTabela(pobjWord, lCount, 5, "0,00", False, wdAlignParagraphRight)
            Call mpCarregarCelulaTabela(pobjWord, lCount, 6, "0,00", False, wdAlignParagraphRight)
            Call mpCarregarCelulaTabela(pobjWord, lCount, 7, "0,00", False, wdAlignParagraphRight)

            lCount = lCount + 1
            Call mpCarregarCelulaTabela(pobjWord, lCount, 1, "SUB TOTAIS", False, wdAlignParagraphRight)
            Call mpCarregarCelulaTabela(pobjWord, lCount, 3, "100,00", False, wdAlignParagraphRight)
            Call mpCarregarCelulaTabela(pobjWord, lCount, 5, "0,00", False, wdAlignParagraphRight)
            Call mpCarregarCelulaTabela(pobjWord, lCount, 6, "0,00", False, wdAlignParagraphRight)
            Call mpCarregarCelulaTabela(pobjWord, lCount, 7, "0,00", False, wdAlignParagraphRight)

            .Range(.Range.End - 1).Select
            .Paragraphs.Add
            .Paragraphs(.Paragraphs.Count).Range.Text = "CONTA DO PRÊMIO R$"
            .Paragraphs(.Paragraphs.Count).Range.bold = True
            .Paragraphs.Add
            .Paragraphs.Add
            .Paragraphs(.Paragraphs.Count).Range.bold = False
            .Paragraphs(.Paragraphs.Count).Range.Text = "Prêmio Líq.total -------" & vbTab & Format(pobjstcPed.VAL_PR, "standard")
            .Paragraphs.Add
            .Paragraphs.Add
            .Paragraphs(.Paragraphs.Count).Range.Text = "Adicional ----------------" & vbTab & "0,00"
            .Paragraphs.Add
            .Paragraphs.Add
            .Paragraphs(.Paragraphs.Count).Range.Text = "Custo apólice ----------" & vbTab & Format(pobjstcPed.Val_Cust_Emis, "standard")
            .Paragraphs.Add
            .Paragraphs.Add
            .Paragraphs(.Paragraphs.Count).Range.Text = "I.O.F. ---------------------" & vbTab & Format(pobjstcPed.Val_IOF, "standard")
            .Paragraphs.Add
            .Paragraphs.Add
            .Paragraphs(.Paragraphs.Count).Range.Text = "Prêmio total-------------" & vbTab & Format(pobjstcPed.VAL_PR + pobjstcPed.Val_Cust_Emis + pobjstcPed.Val_IOF, "standard")
        End If
        'Beneficiário
        If pobjstcPed.ColecaoTextoEspec.Count > 0 Then

            Call mpNovaPaginaOutrosDocs(pobjWord, "BENEFICIÁRIO", lNovaPag)
            lNovaPag = True
            .Paragraphs.Add
            .Paragraphs(.Paragraphs.Count).Range.Text = "CLÁUSULA BENEFICIÁRIA"
            .Paragraphs(.Paragraphs.Count).Range.Underline = True
            .Paragraphs(.Paragraphs.Count).Range.bold = True
            .Paragraphs.Add
            .Paragraphs.Add
            .Paragraphs(.Paragraphs.Count).Range.Underline = wdUnderlineNone
            .Paragraphs(.Paragraphs.Count).Range.bold = False
            .Paragraphs(.Paragraphs.Count).Range.Text = "Fica entendido e acordado que toda e qualquer indenização " & _
                                                        "devida pelo presente Contrato de Seguro referente ao prédio deverá ser paga, em caso de eventual sinistro a (*) vide " & _
                                                        "abaixo, na qualidade de Proprietário do Imóvel Segurado."
            .Paragraphs.Add
            .Paragraphs.Add
            .Paragraphs(.Paragraphs.Count).Range.Text = "(*)"
            .Paragraphs.Add
            .Paragraphs.Add
            .Paragraphs(.Paragraphs.Count).Range.Text = "ITEM NOME"
            .Paragraphs.Add
            lCount = 0
            For Each lobjstcBenef In pobjstcPed.ColecaoTextoEspec
                .Paragraphs.Add
                lCount = lCount + 1
                .Paragraphs(.Paragraphs.Count).Range.Text = Format(lCount, "000") & " - " & lobjstcBenef.Espec_lin1
                If gM46V999.gfPreenchido(lobjstcBenef.Espec_lin2) Then
                    .Paragraphs.Add
                    lCount = lCount + 1
                    .Paragraphs(.Paragraphs.Count).Range.Text = Format(lCount, "000") & " - " & lobjstcBenef.Espec_lin2
                End If
                If gM46V999.gfPreenchido(lobjstcBenef.Espec_lin3) Then
                    .Paragraphs.Add
                    lCount = lCount + 1
                    .Paragraphs(.Paragraphs.Count).Range.Text = Format(lCount, "000") & " - " & lobjstcBenef.Espec_lin3
                End If
                If gM46V999.gfPreenchido(lobjstcBenef.Espec_lin4) Then
                    .Paragraphs.Add
                    lCount = lCount + 1
                    .Paragraphs(.Paragraphs.Count).Range.Text = Format(lCount, "000") & " - " & lobjstcBenef.Espec_lin4
                End If
                If gM46V999.gfPreenchido(lobjstcBenef.Espec_lin5) Then
                    .Paragraphs.Add
                    lCount = lCount + 1
                    .Paragraphs(.Paragraphs.Count).Range.Text = Format(lCount, "000") & " - " & lobjstcBenef.Espec_lin5
                End If
                If gM46V999.gfPreenchido(lobjstcBenef.Espec_lin6) Then
                    .Paragraphs.Add
                    lCount = lCount + 1
                    .Paragraphs(.Paragraphs.Count).Range.Text = Format(lCount, "000") & " - " & lobjstcBenef.Espec_lin6
                End If
                If gM46V999.gfPreenchido(lobjstcBenef.Espec_lin7) Then
                    .Paragraphs.Add
                    lCount = lCount + 1
                    .Paragraphs(.Paragraphs.Count).Range.Text = Format(lCount, "000") & " - " & lobjstcBenef.Espec_lin7
                End If
                If gM46V999.gfPreenchido(lobjstcBenef.Espec_lin8) Then
                    .Paragraphs.Add
                    lCount = lCount + 1
                    .Paragraphs(.Paragraphs.Count).Range.Text = Format(lCount, "000") & " - " & lobjstcBenef.Espec_lin8
                End If
                If gM46V999.gfPreenchido(lobjstcBenef.Espec_lin9) Then
                    .Paragraphs.Add
                    lCount = lCount + 1
                    .Paragraphs(.Paragraphs.Count).Range.Text = Format(lCount, "000") & " - " & lobjstcBenef.Espec_lin9
                End If
                If gM46V999.gfPreenchido(lobjstcBenef.Espec_lin10) Then
                    .Paragraphs.Add
                    lCount = lCount + 1
                    .Paragraphs(.Paragraphs.Count).Range.Text = Format(lCount, "000") & " - " & lobjstcBenef.Espec_lin10
                End If
                If gM46V999.gfPreenchido(lobjstcBenef.Espec_lin11) Then
                    lCount = lCount + 1
                    .Paragraphs.Add
                    .Paragraphs(.Paragraphs.Count).Range.Text = Format(lCount, "000") & " - " & lobjstcBenef.Espec_lin11
                End If
                If gM46V999.gfPreenchido(lobjstcBenef.Espec_lin12) Then
                    .Paragraphs.Add
                    lCount = lCount + 1
                    .Paragraphs(.Paragraphs.Count).Range.Text = Format(lCount, "000") & " - " & lobjstcBenef.Espec_lin12
                End If
            Next lobjstcBenef
        End If


        'Bens
        For Each lobjstcLoc In pobjstcPed.ColecaoPedLoc
            If lobjstcLoc.ColecaoPedLocBem.Count > 0 Then

                Call mpNovaPaginaOutrosDocs(pobjWord, "BENS", lNovaPag)
                lNovaPag = True

                Call mpCriarTabela(pobjWord, lobjstcLoc.ColecaoPedLocBem.Count + 4, 4, True)
                Call mpCarregarCelulaTabela(pobjWord, 1, 1, "Código", True, wdAlignParagraphLeft)
                Call mpCarregarCelulaTabela(pobjWord, 1, 2, "Bem", True, wdAlignParagraphLeft)
                Call mpCarregarCelulaTabela(pobjWord, 1, 3, "Descrição", True, wdAlignParagraphLeft)
                Call mpCarregarCelulaTabela(pobjWord, 1, 4, "Valor", True, wdAlignParagraphLeft)
                lCount = 2
                gM46V999.gCodProduto = pobjstcPed.Cod_Produto
                gM46V999.gCodSequencia = pobjstcPed.Cod_Sequencia
                For Each lObjstcBens In lobjstcLoc.ColecaoPedLocBem
                    lCount = lCount + 1
                    Call mpCarregarCelulaTabela(pobjWord, lCount, 1, Format(lObjstcBens.Cod_Bem, "000"), False, wdAlignParagraphLeft, 1.5)
                    Call mpCarregarCelulaTabela(pobjWord, lCount, 2, lObjstcBens.gTipBem, False, wdAlignParagraphLeft, 5)
                    Call mpCarregarCelulaTabela(pobjWord, lCount, 3, lObjstcBens.Dsc_Bem, False, wdAlignParagraphLeft, 8)
                    Call mpCarregarCelulaTabela(pobjWord, lCount, 4, Format(lObjstcBens.Val_Bem, "standard"), False, wdAlignParagraphRight, 3)
                    lValTotalBem = lValTotalBem + lObjstcBens.Val_Bem
                Next lObjstcBens
                lCount = lCount + 2
                Call mpCarregarCelulaTabela(pobjWord, lCount, 3, "TOTAL", False, wdAlignParagraphRight)
                Call mpCarregarCelulaTabela(pobjWord, lCount, 4, Format(lValTotalBem, "standard"), False, wdAlignParagraphRight)
            End If
        Next lobjstcLoc

        lNomeArqFrac = mLetraMapeamento & ":\DOCDIVERSOS\" & Format(pobjstcPed.Num_Apol, "0000000000") & Format(pobjstcPed.Num_Endo, "000000") & "FRAC.txt"
        If Dir(lNomeArqFrac) <> "" Then
            Call mpNovaPaginaOutrosDocs(pobjWord, "CLÁUSULA DE FRACIONAMENTO DO PRÊMIO", lNovaPag)
            lNovaPag = True
            lAuxiliar = "1) Fica entendido que o Prêmio Líquido da Apólice ou do Aditivo " & _
                        "será pago em " & pobjstcPed.Qtd_Parc & " parcelas mensais e sucessivas, " & _
                        "vencendo-se a primeira em " & pobjstcPed.mAuxPriVenc & " , de " & _
                        "conformidade com o disposto no quadro a seguir:"
            .Paragraphs.Add
            .Paragraphs(.Paragraphs.Count).Range.Text = lAuxiliar
            .Paragraphs.Add
            .Paragraphs.Add
            Call mpCriarTabela(pobjWord, pobjstcPed.Qtd_Parc + 2, 7, False)

            Call mpCarregarCelulaTabela(pobjWord, 1, 1, "Parc.", False, wdAlignParagraphCenter)
            Call mpCarregarCelulaTabela(pobjWord, 1, 2, "Prêmio Líquido (R$)", False, wdAlignParagraphCenter)
            Call mpCarregarCelulaTabela(pobjWord, 1, 3, "Juros de Fracionamento (R$)", False, wdAlignParagraphCenter)
            Call mpCarregarCelulaTabela(pobjWord, 1, 4, "Custo de Apólice", False, wdAlignParagraphCenter)
            Call mpCarregarCelulaTabela(pobjWord, 1, 5, "I.O.F. (R$)", False, wdAlignParagraphCenter)
            Call mpCarregarCelulaTabela(pobjWord, 1, 6, "Prêmio Total (R$)", False, wdAlignParagraphCenter)
            Call mpCarregarCelulaTabela(pobjWord, 1, 7, "Limite Pgto", False, wdAlignParagraphCenter)

            lArquivo = FreeFile
            Open lNomeArqFrac For Input Access Read As #lArquivo
            lLinhas = 0
            lCount = 1
            While Not EOF(lArquivo)
                Line Input #lArquivo, lRegistro
                lLinhas = lLinhas + 1

                Select Case lLinhas
                Case 12 To 23
                    If gM46V999.gfPreenchido(lRegistro) Then
                        lCount = lCount + 1
                        Call mpCarregarCelulaTabela(pobjWord, lCount, 1, Trim(Mid(lRegistro, 1, 5)), False, wdAlignParagraphCenter)
                        Call mpCarregarCelulaTabela(pobjWord, lCount, 2, Trim(Mid(lRegistro, 6, 20)), False, wdAlignParagraphCenter)
                        Call mpCarregarCelulaTabela(pobjWord, lCount, 3, Trim(Mid(lRegistro, 26, 19)), False, wdAlignParagraphCenter)
                        Call mpCarregarCelulaTabela(pobjWord, lCount, 4, Trim(Mid(lRegistro, 45, 13)), False, wdAlignParagraphCenter)
                        Call mpCarregarCelulaTabela(pobjWord, lCount, 5, Trim(Mid(lRegistro, 58, 13)), False, wdAlignParagraphCenter)
                        Call mpCarregarCelulaTabela(pobjWord, lCount, 6, Trim(Mid(lRegistro, 71, 17)), False, wdAlignParagraphCenter)
                        Call mpCarregarCelulaTabela(pobjWord, lCount, 7, Trim(Mid(lRegistro, 88, 11)), False, wdAlignParagraphCenter)
                    End If
                Case 24
                    lCount = lCount + 1
                    Call mpCarregarCelulaTabela(pobjWord, lCount, 1, "", False, wdAlignParagraphCenter)
                    Call mpCarregarCelulaTabela(pobjWord, lCount, 2, Trim(Mid(lRegistro, 2, 26)), False, wdAlignParagraphCenter)
                    Call mpCarregarCelulaTabela(pobjWord, lCount, 3, Trim(Mid(lRegistro, 27, 18)), False, wdAlignParagraphCenter)
                    Call mpCarregarCelulaTabela(pobjWord, lCount, 4, Trim(Mid(lRegistro, 45, 13)), False, wdAlignParagraphCenter)
                    Call mpCarregarCelulaTabela(pobjWord, lCount, 5, Trim(Mid(lRegistro, 58, 13)), False, wdAlignParagraphCenter)
                    Call mpCarregarCelulaTabela(pobjWord, lCount, 6, Trim(Mid(lRegistro, 71, 17)), False, wdAlignParagraphCenter)
                    Call mpCarregarCelulaTabela(pobjWord, lCount, 7, "", False, wdAlignParagraphCenter)
                End Select
            Wend
            Close #lArquivo

            .Paragraphs.Add
            .Paragraphs.Add

            lAuxiliar = "2) Taxa efetiva de juros pactuada: " & pobjstcPed.Val_Juro & "% am" & vbCrLf & vbCrLf & _
                        "3) Fica garantido ao Segurado a liquidação antecipada do prêmio fracionado, total ou parcialmente, " & _
                        "mediante redução proporcional dos  juros contratados." & vbCrLf & vbCrLf & _
                        "4) A falta de pagamento de qualquer parcela do fracionamento, após transcorridos os " & _
                        "prazos referidos nas Condições Gerais da Apólice e os estabelecidos pela Tabela de Prazo " & _
                        "curto constante na Circular nº 239, de 22.12.03 da superintendência de Seguros Privados - SUSEP, " & _
                        "implicará no cancelamento automático e de pleno direito do respectivo contrato de seguro " & _
                        "ou seu aditamento, independente de qualquer interpelação judicial ou extrajudicial." & vbCrLf & vbCrLf & _
                        "5) O seguro poderá restabelecer o direito às coberturas contratadas, pelo período inicialmente" & _
                        "acordado, desde que retorne o pagamento do prêmio devido, dentro do prazo estabelecido no parágrafo " & _
                        "anterior, mediante pagamento do prêmio acrescido dos juros de mora de 4,5% a m. (pro-rata-die)." & vbCrLf & vbCrLf & _
                        "6) O direito a indenização não ficará prejudicado quando o sinistro ocorrer dentro do prazo de " & _
                        "fracionamento do prêmio." & vbCrLf & vbCrLf & _
                        "7) Serão descontadas do valor da indenização as parcelas ainda não pagas, " & _
                        "reduzidos proporcionalmente os juros pactuados, se o sinistro acarretar a indenização integral."

            .Paragraphs.Add
            .Paragraphs(.Paragraphs.Count).Range.Text = lAuxiliar
        End If


        For Each lobjstcLoc In pobjstcPed.ColecaoPedLoc
            If lobjstcLoc.ColecaoPedSeguro.Count > 0 Then
                Call mpNovaPaginaOutrosDocs(pobjWord, "OUTROS SEGUROS", lNovaPag)
                lNovaPag = True
                lAuxiliar = " Existem outros seguros garantindo os mesmos bens segurados pela presente apólice."
                .Paragraphs(.Paragraphs.Count).Range.Text = lAuxiliar

                Call mpCriarTabela(pobjWord, lobjstcLoc.ColecaoPedSeguro.Count + 2, 4, True)
                Call mpCarregarCelulaTabela(pobjWord, 1, 1, "Seguradora", True, wdAlignParagraphLeft, 10)
                Call mpCarregarCelulaTabela(pobjWord, 1, 2, "Apólice", True, wdAlignParagraphRight, 3)
                Call mpCarregarCelulaTabela(pobjWord, 1, 3, "Item", True, wdAlignParagraphRight, 3)
                Call mpCarregarCelulaTabela(pobjWord, 1, 4, "Sinistro?", True, wdAlignParagraphRight, 3)
                lCount = 2
                For Each lobjstcSeguro In lobjstcLoc.ColecaoPedSeguro
                    lCount = lCount + 1
                    Call mpCarregarCelulaTabela(pobjWord, lCount, 1, lobjstcSeguro.Dsc_Outro_Seguro, False, wdAlignParagraphLeft)
                    Call mpCarregarCelulaTabela(pobjWord, lCount, 2, lobjstcSeguro.Num_Apol_Ant, False, wdAlignParagraphRight)
                    Call mpCarregarCelulaTabela(pobjWord, lCount, 3, lobjstcSeguro.Num_Item_Ant, False, wdAlignParagraphRight)
                    Call mpCarregarCelulaTabela(pobjWord, lCount, 4, IIf(lobjstcSeguro.Ind_Sinistro = e_SimNao_Sim, "Sim", "Não"), False, wdAlignParagraphRight)
                Next lobjstcSeguro
            End If
        Next lobjstcLoc
    End With
    mfGerarDemaisDocs = True
End Function

Private Function mfGerarArqApolice(pNomeArquivo As String, pobjstcPed As stcA46V702B, _
                                   pobjWord As Word.Application) As Boolean
    Dim objDoc As Word.Document
    Dim lPesqDigito As String
    Dim clsA99V001 As clsA99V001
    Dim lInd As Integer
    Dim lMensagem As String
    Dim lCNPJForm As String
    Dim lObjstcCorr As stcA46V705B
    Dim lPremioTotal As Double
    Dim lAuxiliar As String
    Dim lobjstcLoc As stcA46V708B
    Dim lValorParc As Variant
    Dim lObjstcVenc As stcA46V713B
    Dim lI As Long
    Dim lParcelas As Variant

    'If gfForNormalDot(pobjWord, lMensagem) = False Then
    '    GoTo ControleErro
    'End If

    On Error GoTo ControleErro

    pobjWord.Application.DisplayAlerts = wdAlertsNone

    If Dir(gM46V999.gPathArquivos & "apolice.doc") = "" Then
        Call gM46V999.gpGraLog(1, "Arquivo layout da apólice não foi localizado (" & gM46V999.gPathArquivos & "apolice.doc" & ")")
        Exit Function
    End If
    Call FileCopy(gM46V999.gPathArquivos & "apolice.doc", gM46V999.gAppPath & "apolice.doc")
    Set objDoc = pobjWord.Documents.Open(gM46V999.gAppPath & "apolice.doc")

    'pobjWord.Visible = False


    Set objDoc = pobjWord.Documents(1)
    lAuxiliar = pobjstcPed.gCodSusep
    Call mpReplaceWord(objDoc, "@@NUM_SUSEP@@", lAuxiliar)
    Call mpReplaceWord(objDoc, "@@RAMO@@", pobjstcPed.gDesc_Ramo)
    Call mpReplaceWord(objDoc, "@@ORGAO_EMISSOR@@", "SÃO PAULO")
    Call mpReplaceWord(objDoc, "@@UNIDADE@@", Format(pobjstcPed.Cod_Unid_Prod, "0000"))


    Set clsA99V001 = New clsA99V001
    Set clsA99V001.M46V999 = gM46V999

    Call clsA99V001.DIG_YASUDA(pobjstcPed.Num_Apol, lPesqDigito, lMensagem)
    pobjstcPed.mDigitApolice = lPesqDigito
    Call mpReplaceWord(objDoc, "@@APOLICE@@", Format(pobjstcPed.Num_Apol, "0000000000") & "-" & lPesqDigito)
    Call mpReplaceWord(objDoc, "@@REN_APOLICE@@", Format(pobjstcPed.objstcCtrlEmis.Num_Apol_Ant, "0000000000"))
    Call mpReplaceWord(objDoc, "@@PROPOSTA@@", Format(pobjstcPed.objstcCtrlEmis.Num_Prop, "000000"))
    Call mpReplaceWord(objDoc, "@@SEGURADO@@", pobjstcPed.Nom_Segurado)
    Call mpReplaceWord(objDoc, "@@CPF_CNPJ@@", pobjstcPed.gCNPJ_CPF_Formatado)
    Call mpReplaceWord(objDoc, "@@ENDERECO@@", Trim(pobjstcPed.Cod_Tip_Logr) & " " & Trim(pobjstcPed.Nom_Logr) & ", " & Trim(pobjstcPed.Dsc_Compl_Logr) & " " & Trim(pobjstcPed.Nom_Bairro))

    'Corretores
    Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(1)
    Call mpReplaceWord(objDoc, "@@CORRETOR_1@@", lObjstcCorr.gNom_Corretor)
    Call mpReplaceWord(objDoc, "@@SUSEP_CORR_1@@", lObjstcCorr.mNumSusep)

    If pobjstcPed.ColecaoPedCorr.Count > 1 Then
        Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(2)
    Else
        Set lObjstcCorr = New stcA46V705B
        Set lObjstcCorr.M46V999 = gM46V999

    End If

    Call mpReplaceWord(objDoc, "@@CORRETOR_2@@", lObjstcCorr.gNom_Corretor)
    Call mpReplaceWord(objDoc, "@@SUSEP_CORR_2@@", lObjstcCorr.mNumSusep)

    If pobjstcPed.ColecaoPedCorr.Count > 2 Then
        Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(3)
    Else
        Set lObjstcCorr = New stcA46V705B
        Set lObjstcCorr.M46V999 = gM46V999

    End If

    Call mpReplaceWord(objDoc, "@@CORRETOR_3@@", lObjstcCorr.gNom_Corretor)
    Call mpReplaceWord(objDoc, "@@SUSEP_CORR_3@@", lObjstcCorr.mNumSusep)

    lAuxiliar = Format(pobjstcPed.Val_IS, "standard")
    Call mpReplaceWord(objDoc, "@@LMG@@", lAuxiliar)

    lAuxiliar = Format(pobjstcPed.VAL_PR, "standard")
    Call mpReplaceWord(objDoc, "@@PREMIO_TARIF@@", lAuxiliar)
    Call mpReplaceWord(objDoc, "@@DESCONTO@@", "-")

    lAuxiliar = Format(pobjstcPed.VAL_PR, "standard")
    Call mpReplaceWord(objDoc, "@@PREMIO_LIQUIDO@@", lAuxiliar)

    lAuxiliar = Format(pobjstcPed.Val_Juro, "standard")
    Call mpReplaceWord(objDoc, "@@JUROS@@", lAuxiliar)

    Call mpReplaceWord(objDoc, "@@TX@@", "-")

    lAuxiliar = Format(0, "standard")
    Call mpReplaceWord(objDoc, "@@PERCAM@@", lAuxiliar)

    lAuxiliar = Format(pobjstcPed.Val_Cust_Emis, "standard")
    Call mpReplaceWord(objDoc, "@@CUSTO_EMISSAO@@", lAuxiliar)

    lAuxiliar = Format(pobjstcPed.Val_IOF, "standard")
    Call mpReplaceWord(objDoc, "@@IOF@@", lAuxiliar)

    lPremioTotal = pobjstcPed.VAL_PR + pobjstcPed.Val_Cust_Emis + pobjstcPed.Val_IOF
    lAuxiliar = Format(lPremioTotal, "standard")
    Call mpReplaceWord(objDoc, "@@PREMIO_TOTAL@@", lAuxiliar)

    Set lobjstcLoc = pobjstcPed.ColecaoPedLoc.Item(1)
    'Calcula o fracionamento
    With pobjstcPed
        Call gM46V998.gfCalFracionamento(pobjstcPed, lobjstcLoc.Tip_Emissao, _
                                         lValorParc, lMensagem, 7, pobjstcPed.Cod_Reg, pobjstcPed.Val_IOF, .Val_Parc_Min)
    End With
    If Not mfCarregarVencimento(pobjstcPed, lParcelas) Then GoTo ControleErro


    For lI = 0 To UBound(lParcelas) - 1
        lAuxiliar = lParcelas(lI)
        If lI = 0 Then pobjstcPed.mAuxPriVenc = lAuxiliar
        Call mpReplaceWord(objDoc, "@@PARCELA_" & lI + 1 & "@@", lAuxiliar)
    Next lI
    For lI = UBound(lParcelas) To 9
        Call mpReplaceWord(objDoc, "@@PARCELA_" & lI + 1 & "@@", "")
    Next lI

    'Vigências
    Call mpReplaceWord(objDoc, "@@INICIO_VIGENCIA@@", "24 Hs " & UCase(Format(gM46V999.gfForData(2, pobjstcPed.Dat_Ini_Vig), "dd mmm yyyy")))
    Call mpReplaceWord(objDoc, "@@TERMINO_VIGENCIA@@", "24 Hs " & UCase(Format(gM46V999.gfForData(2, pobjstcPed.Dat_Ter_Vig), "dd mmm yyyy")))

    Call mpReplaceWord(objDoc, "@@LOCAL_DATA@@", "SÃO PAULO " & Format(Date, "dd/mm/yyyy"))
    Call mpReplaceWord(objDoc, "@@APOLICE_RODAPE@@", "Apólice " & Format(pobjstcPed.Num_Apol, "0000000000") & "-" & lPesqDigito, True)

    '@@ESPECIFICACAO@@
    objDoc.SaveAs pNomeArquivo
    'Imprimir:
    '    If Not mfImprimir(pobjWord, pObjstcPed, lMensagem) Then GoTo ControleErro

Finalizar:
    mfGerarArqApolice = True
    Exit Function
ControleErro:
    If Err.Number <> 0 Then
        Call gM46V999.gpGraLog(1, Err.Number & " - " & Err.Description)
    End If
    If gM46V999.gfPreenchido(lMensagem) = True Then
        Call gM46V999.gpGraLog(1, "mfGerarArqApolice " & Chr(13) & lMensagem)
    End If
    pobjWord.Quit wdDoNotSaveChanges
    Set pobjWord = Nothing
    Set objDoc = Nothing

End Function

Private Function mfCarregarVencimento(pobjstcPed As stcA46V702B, ByRef pParcelas As Variant) As Boolean
    Dim lNomeArqFrac As String
    Dim lArquivo As Integer
    Dim lLinhas As Long
    Dim lCount As Long
    Dim lRegistro As String
    Dim lAuxiliar As String
    lNomeArqFrac = mLetraMapeamento & ":\DOCDIVERSOS\" & Format(pobjstcPed.Num_Apol, "0000000000") & Format(pobjstcPed.Num_Endo, "000000") & "FRAC.txt"
    If Dir(lNomeArqFrac) <> "" Then
        lArquivo = FreeFile
        Open lNomeArqFrac For Input Access Read As #lArquivo
        lLinhas = 0
        lCount = 1
        While Not EOF(lArquivo)
            Line Input #lArquivo, lRegistro
            lLinhas = lLinhas + 1
            Select Case lLinhas
            Case 12 To 23
                If gM46V999.gfPreenchido(lRegistro) Then
                    lCount = lCount + 1
                    If gM46V999.gfPreenchido(lAuxiliar) Then
                        lAuxiliar = lAuxiliar & ";"
                    End If
                    lAuxiliar = lAuxiliar & Mid(lRegistro, 88, 11)
                End If
            End Select
        Wend
        Close #lArquivo
    End If
    lAuxiliar = lAuxiliar & ";"
    pParcelas = Split(lAuxiliar, ";")
    mfCarregarVencimento = True
End Function
Private Function mfSeparador(pobjstc As stcA46V702B, pQtde As Integer) As Boolean
    Dim objDoc As Word.Document
    Dim lMensagem As String
    Dim lAuxiliar As String
    Dim lobjWord As Word.Application


    On Error GoTo ControleErro
    Set lobjWord = New Word.Application
    lobjWord.Application.DisplayAlerts = wdAlertsNone

    If pQtde = 1 Then
        Call FileCopy(gM46V999.gPathArquivos & "SeparadorSeg.doc", gM46V999.gAppPath & "SeparadorSeg.doc")
        Set objDoc = lobjWord.Documents.Open(gM46V999.gAppPath & "SeparadorSeg.doc")
    ElseIf pQtde = 2 Then
        Call FileCopy(gM46V999.gPathArquivos & "SeparadorCorr.doc", gM46V999.gAppPath & "SeparadorCorr.doc")
        Set objDoc = lobjWord.Documents.Open(gM46V999.gAppPath & "SeparadorCorr.doc")
    Else
        Call FileCopy(gM46V999.gPathArquivos & "SeparadorFinan.doc", gM46V999.gAppPath & "SeparadorFinan.doc")
        Set objDoc = lobjWord.Documents.Open(gM46V999.gAppPath & "SeparadorFinan.doc")
    End If

    'lobjWord.Visible = False

    'Call mpReplaceWord(objDoc, "@@Apolice@@", Format(objstcCtrlEmissao.Num_Apol, "0000000000"))
    'Call mpReplaceWord(objDoc, "@@Endosso@@", Format(objstcCtrlEmissao.Num_Endo, "000000"))

    If Not mfImprimir(lobjWord, pobjstc, lMensagem, False) Then GoTo ControleErro
Fim:
    lobjWord.Application.Quit wdDoNotSaveChanges

    mfSeparador = True
    Exit Function
ControleErro:
    If Err.Number <> 0 Then
        Call gM46V999.gpGraLog(1, Err.Number & " - " & Err.Description)
    End If
    If gM46V999.gfPreenchido(lMensagem) = True Then
        Call gM46V999.gpGraLog(1, "mfSeparador " & Chr(13) & lMensagem)
    End If
    lobjWord.Quit wdDoNotSaveChanges
    Set lobjWord = Nothing
    Set objDoc = Nothing

End Function
Private Function mfImprimir(ByRef pobjWord As Word.Application, pobjstc As stcA46V702B, _
                            ByRef pMensagem As String, ByVal pFrenteVerso As Boolean) As Boolean
    Dim lTimeInicial As String
    Dim lUltImpressora As String
    Dim lImpressora As String
    On Error GoTo ControleErro

    lUltImpressora = pobjWord.ActivePrinter

    'Imprimir
    If pFrenteVerso = True Then
        lImpressora = gM46V999.gImprFrenteVerso
    Else
        lImpressora = gM46V999.gImprFrente
    End If
    If gM46V999.gfPreenchido(lImpressora) Then
        pobjWord.ActivePrinter = lImpressora
        pobjWord.PrintOut
    End If

    lTimeInicial = Now
    Do While pobjWord.BackgroundPrintingStatus <> 0
        panAviso.Caption = "Aguardando a fila de impressão (" & pobjWord.BackgroundPrintingStatus & ")... " & Format(DateDiff("s", lTimeInicial, Now) \ 60, "00") & ":" & Format(DateDiff("s", lTimeInicial, Now) Mod 60, "00")
        DoEvents
        Call gM46V111.gpPausa(0.5)
    Loop
    panAviso.Caption = "Gerando arquivos da apólices " & Format(pobjstc.Num_Apol, "0000000000") & " ( PI : " & pobjstc.Num_PI & ")"
    pobjWord.ActivePrinter = lUltImpressora

    mfImprimir = True
    Exit Function
ControleErro:
    If Err.Number = 5216 Then
        pMensagem = "Impressora configurada no arquivo INI está errada, ela não existe (" & lImpressora & ")."
    Else
        pMensagem = "mfImprimir " & Err.Description
    End If
End Function
Private Function mfGerarArqEndosso(pNomeArquivo As String, pobjstcPed As stcA46V702B, _
                                   pobjWord As Word.Application) As Boolean
    Dim objDoc As Word.Document
    Dim lPesqDigito As String
    Dim clsA99V001 As clsA99V001
    Dim lMensagem As String
    Dim lCNPJForm As String
    Dim lPremioTotal As Double
    Dim lAuxiliar As String
    Dim lUltImpressora As String
    Dim lObjstcCorr As stcA46V705B
    Dim lobjstcDizer As stcA46V722B


    'If gfForNormalDot(pobjWord, lMensagem) = False Then
    '    GoTo ControleErro
    'End If

    On Error GoTo ControleErro

    pobjWord.Application.DisplayAlerts = wdAlertsNone

    If Dir(gM46V999.gPathArquivos & "endosso.doc") = "" Then
        Call gM46V999.gpGraLog(1, "Arquivo layout do endosso não foi localizado (" & gM46V999.gPathArquivos & "endosso.doc" & ")")
        Exit Function
    End If
    Call FileCopy(gM46V999.gPathArquivos & "endosso.doc", gM46V999.gAppPath & "endosso.doc")
    Set objDoc = pobjWord.Documents.Open(gM46V999.gAppPath & "endosso.doc")
    'pobjWord.Visible = False

    Call mpReplaceWord(objDoc, "@@NUM_SUSEP@@", pobjstcPed.gCodSusep)
    Call mpReplaceWord(objDoc, "@@RAMO@@", pobjstcPed.gDesc_Ramo)
    Call mpReplaceWord(objDoc, "@@ORGAO_EMISSOR@@", "SÃO PAULO")
    Call mpReplaceWord(objDoc, "@@UNIDADE@@", Format(pobjstcPed.Cod_Unid_Prod, "0000"))
    Call mpReplaceWord(objDoc, "@@PRODUTOR@@", Format(pobjstcPed.Cod_Prod, "0000"))

    Set clsA99V001 = New clsA99V001
    Set clsA99V001.M46V999 = gM46V999

    Call clsA99V001.DIG_YASUDA(pobjstcPed.objstcCtrlEmis.Num_Apol, lPesqDigito, lMensagem)
    Call mpReplaceWord(objDoc, "@@APOLICE@@", Format(pobjstcPed.objstcCtrlEmis.Num_Apol, "0000000000") & "-" & lPesqDigito)
    Call mpReplaceWord(objDoc, "@@ENDOSSO@@", Format(pobjstcPed.objstcCtrlEmis.Num_Endo, "000000"))

    Call mpReplaceWord(objDoc, "@@SEGURADO@@", pobjstcPed.Nom_Segurado)
    Call mpReplaceWord(objDoc, "@@CPF_CNPJ@@", pobjstcPed.gCNPJ_CPF_Formatado)

    'Corretores
    Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(1)
    Call mpReplaceWord(objDoc, "@@CORRETOR_1@@", lObjstcCorr.gNom_Corretor)
    Call mpReplaceWord(objDoc, "@@SUSEP_CORR_1@@", lObjstcCorr.mNumSusep)

    If pobjstcPed.ColecaoPedCorr.Count > 1 Then
        Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(2)
    Else
        Set lObjstcCorr = New stcA46V705B
        Set lObjstcCorr.M46V999 = gM46V999

    End If

    Call mpReplaceWord(objDoc, "@@CORRETOR_2@@", lObjstcCorr.gNom_Corretor)
    Call mpReplaceWord(objDoc, "@@SUSEP_CORR_2@@", lObjstcCorr.mNumSusep)

    If pobjstcPed.ColecaoPedCorr.Count > 2 Then
        Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(3)
    Else
        Set lObjstcCorr = New stcA46V705B
        Set lObjstcCorr.M46V999 = gM46V999

    End If

    Call mpReplaceWord(objDoc, "@@CORRETOR_3@@", lObjstcCorr.gNom_Corretor)
    Call mpReplaceWord(objDoc, "@@SUSEP_CORR_3@@", lObjstcCorr.mNumSusep)

    'Vigências
    Call mpReplaceWord(objDoc, "@@INICIO_VIGENCIA@@", "24 Hs " & UCase(Format(gM46V999.gfForData(2, pobjstcPed.Dat_Ini_Vig), "dd mmm yyyy")))
    Call mpReplaceWord(objDoc, "@@TERMINO_VIGENCIA@@", "24 Hs " & UCase(Format(gM46V999.gfForData(2, pobjstcPed.Dat_Ter_Vig), "dd mmm yyyy")))

    Call mpReplaceWord(objDoc, "@@PRAZO_PAGTO@@", "")

    lAuxiliar = Format(pobjstcPed.VAL_PR, "standard")
    Call mpReplaceWord(objDoc, "@@PREMIO_LIQUIDO@@", lAuxiliar)

    Call mpReplaceWord(objDoc, "@@ADICIONAIS@@", "")

    lAuxiliar = Format(pobjstcPed.Val_Cust_Emis, "standard")
    Call mpReplaceWord(objDoc, "@@CUSTO_EMISSAO@@", lAuxiliar)

    lAuxiliar = Format(pobjstcPed.Val_IOF, "standard")
    Call mpReplaceWord(objDoc, "@@IOF@@", lAuxiliar)


    lPremioTotal = pobjstcPed.VAL_PR + pobjstcPed.Val_Cust_Emis + pobjstcPed.Val_IOF
    lAuxiliar = Format(lPremioTotal, "standard")
    Call mpReplaceWord(objDoc, "@@PREMIO_TOTAL@@", lAuxiliar)

    lAuxiliar = ""
    For Each lobjstcDizer In pobjstcPed.ColecaoTextoEndo
        If InStr(1, lobjstcDizer.Dsc_Dizer_Endo, vbCrLf) = 0 Then
            lAuxiliar = lAuxiliar & lobjstcDizer.Dsc_Dizer_Endo & vbCrLf
        Else
            lAuxiliar = lAuxiliar & lobjstcDizer.Dsc_Dizer_Endo
        End If
    Next lobjstcDizer
    Call mpReplaceWord(objDoc, "@@DIZER_ENDOSSO@@", lAuxiliar)

    Call mpReplaceWord(objDoc, "@@LOCAL_DATA@@", "SÃO PAULO " & Format(Date, "dd/mm/yyyy"))

    Call mpReplaceWord(objDoc, "@@APOLICE_RODAPE@@", Format(pobjstcPed.objstcCtrlEmis.Num_Apol, "0000000000") & "-" & lPesqDigito)
    Call mpReplaceWord(objDoc, "@@ENDOSSO_RODAPE@@", Format(pobjstcPed.objstcCtrlEmis.Num_Endo, "000000"))


    '@@ESPECIFICACAO@@



    objDoc.SaveAs pNomeArquivo
    'Imprimir:
    '    If Not mfImprimir(pobjWord, pObjstcPed, lMensagem) Then GoTo ControleErro
    '    pobjWord.Application.Quit
    '
    '    'Restaura o estilo do usuário
    '    If gfForNormalDot(pobjWord, lMensagem, True) = False Then
    '        GoTo ControleErro
    '    End If
Finalizar:
    '    Set pobjWord = Nothing
    '    Set objDoc = Nothing
    mfGerarArqEndosso = True
    Exit Function
ControleErro:
    If Err.Number <> 0 Then
        Call gM46V999.gpGraLog(1, Err.Number & " - " & Err.Description)
    End If
    If gM46V999.gfPreenchido(lMensagem) = True Then
        Call gM46V999.gpGraLog(1, "mfGerarArqEndosso " & Chr(13) & lMensagem)
    End If
    pobjWord.Quit wdDoNotSaveChanges
    Set pobjWord = Nothing
    Set objDoc = Nothing

End Function

Private Sub mpCarregarImpressos()
    Dim lDescricao As String
    Dim lExisteArq As Boolean
    Dim lOkImpressao As Boolean
    Dim lQtdeArqFicha As Long
    Dim lQtdeFluxo As Long
    Dim lobjPedEspec As clsA46V721A
    Dim lobjstcPedEspec As stcA46V721B
    Dim lObjstcPed As stcA46V702B

    panAviso.Caption = "Pesquisando apólices e endossos..."

    Set lobjPedEspec = New clsA46V721A
    Set lobjPedEspec.M46V999 = gM46V999

    If Not lobjPedEspec.mfCarregarColecaoImpressao("", "") Then
        Call gM46V999.gpGraLog(1, lobjPedEspec.UltimoErro)
        Exit Sub
    End If
    If lobjPedEspec.Colecao.Count = 0 Then
        Call gM46V999.gpGraLog(1, "Emissão diversos - " & Now & vbCrLf & vbCrLf & "Não existem registros pendentes de emissão.")
        mErro = 1
        Exit Sub
    End If
    If gM46V999.gfPreenchido(gM46V999.gPathArquivos) = False Then
        Call gM46V999.gpGraLog(1, "Verificar arquivo INI. diretório de arquivos de layout não preenchido")
        mErro = 2
        Exit Sub
    End If

    Call mpRetornaLetraParaMapear
    If Not gM46V999.gfPreenchido(mLetraMapeamento) Then Exit Sub
    Call gM46V999.gpGraLog(0, "Mapeado na letra " & mLetraMapeamento)

    Set lObjstcPed = New stcA46V702B
    Set lObjstcPed.M46V999 = gM46V999

    For Each lobjstcPedEspec In lobjPedEspec.Colecao
        If gM46V999.gfPreenchido(lObjstcPed.NossoNumeroIn) Then
            lObjstcPed.NossoNumeroIn = lObjstcPed.NossoNumeroIn & ","
        End If
        lObjstcPed.NossoNumeroIn = lObjstcPed.NossoNumeroIn & "'" & lobjstcPedEspec.NOSSO_NUMERO & "'"
    Next lobjstcPedEspec


    Set mobjPed = New clsA46V702A
    Set mobjPed.M46V999 = gM46V999

    If Not mobjPed.mfCarregarColecao(lObjstcPed, "", "", True, , True) Then
        Call gM46V999.gpGraLog(1, mobjPed.UltimoErro)
        mErro = 3
        Exit Sub
    End If

    cmdOk.Enabled = False

    lstLista.Clear
    lOkImpressao = True
    For Each lObjstcPed In mobjPed.Colecao
        lDescricao = lObjstcPed.Num_PI & " - " & lObjstcPed.Num_Apol & IIf(lObjstcPed.Num_Endo > 0, "/" & lObjstcPed.Num_Endo, "")
        lstLista.AddItem lDescricao
        Call gM46V999.gpGraLog(0, lDescricao)
        lstLista.ItemData(lstLista.NewIndex) = lObjstcPed.Num_PI
        lstLista.ListIndex = lstLista.ListCount - 1
        lstLista.Selected(lstLista.ListIndex) = True
    Next lObjstcPed

    panAviso.Caption = "Total de registros pendentes: " & mobjPed.Colecao.Count
    panAviso.Visible = True
    cmdOk.Enabled = True
End Sub
Private Sub chkSel_Click()
    Dim lI As Long

    For lI = 0 To lstLista.ListCount - 1
        If chkSel.Value = vbChecked Then
            lstLista.Selected(lI) = True
        Else
            lstLista.Selected(lI) = False
        End If
    Next lI
End Sub



Private Sub cmdOK_Click()
    If MsgBox("Confirma a geração e impressão das apólices/endossos de diversos selecionadas?", _
              vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
        If gM46V999.gfPreenchido(gM46V999.gImprFrente) Then
            If MsgBox("OPERADOR, as impressoras '" & gM46V999.gImprFrente & "' e '" & gM46V999.gImprFrenteVerso & "' estão ligadas?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                Exit Sub
            End If
        End If
        Screen.MousePointer = vbHourglass
        Call mpEmitirDocumentos
        Screen.MousePointer = vbDefault
        Unload Me
    End If
End Sub

Private Sub Form_Activate()
    If mErro <> 0 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()

    lConfig = False
    Call gM46V999.gpCenForm(Me)
    mErro = 0
    If mTipoDocumento = 0 Then
        Call mpCarregarImpressos
    End If

End Sub
Private Function mfVisualizarClausulas(pArqApolice As String, pArqEspecif As String, pobjstcPed As stcA46V702B) As Boolean
    Dim Aux As String
    Dim lobjstcCobert As stcA46V704B
    Dim lMensagem As String
    Dim lUltImpressora As String
    Dim lNomeFinal As String
    Dim lI As Long
    Dim lQtdePaginas As Long
    Dim lNomePDF As String
    Dim lFinalArq As Boolean
    Dim lCoberturaRc As Boolean
    Dim lDirClausulas As String
    Dim lTimeInicial As String
    Dim lLocal As String
    Dim lobjstcLoc As stcA46V708B
    Dim lSelect As String
    Dim lrst As clsYasRecordSet
    Dim lConexao As ADODB.Connection

    On Error Resume Next
    Kill gM46V999.gPathArquivos & "PDFCreator, Job *.pdf"
    On Error GoTo ControleErro

    lLocal = 1
    lNomeFinal = gM46V999.gPathArquivos & _
                 Format(pobjstcPed.Num_Apol, "0000000000") & _
                 Format(pobjstcPed.Num_Endo, "000000") & ".doc"

    FileCopy pArqApolice, lNomeFinal

    lNomeFinal = gM46V999.gPathArquivos & _
                 Format(pobjstcPed.Num_Apol, "0000000000") & _
                 Format(pobjstcPed.Num_Endo, "000000") & "_ESP.doc"

    'FileCopy pArqEspecif, lNomeFinal

    lLocal = "2"
    lDirClausulas = gM46V999.gPathArquivos & "Clausulas\"

    lSelect = " select clau.* from siscota.dbo.tab_cotacao_clausula as clau"
    lSelect = lSelect & " where clau.ccotacao = " & gM46V999.gfFormatarCamposSQL(pobjstcPed.objstcCtrlEmis.Num_Cotac)

    If Not gM46V999.gfAbrBasDados("GED", gM46V999.gTipGed, gM46V999.gEndGed, lConexao, lMensagem) Then
        Call gM46V999.gpGraLog(1, lMensagem)
        Exit Function
    End If
    If Not gM46V999.gfObtRegistro(lConexao, lSelect, lrst, lMensagem) Then
        Call gM46V999.gpGraLog(1, lMensagem)
        Exit Function
    End If

    lLocal = "7"
    'Imprime Condição Geral
    Aux = lDirClausulas & "T000000.doc"
    'If Not mfImprimirClausula(Aux, pObjstcPed, lMensagem) Then GoTo ControleErro
    lLocal = "8"

    'Imprime demais cláusulas
    lI = 0
    Do While Not lrst.EOF
        lI = lI + 1
        lLocal = "9 " & lI
        Aux = lDirClausulas & lrst("Cod_Ramo") & Format(lrst("Cod_Clausula"), "0000") & ".doc"
        lLocal = "10 " & lI
        If Dir(Aux) = "" Then
            Call gM46V999.gpGraLog(1, "Documento da cláusula " & lrst("Cod_Clausula") & "(" & Aux & ") não localizado.")
        Else
            If Not mfImprimirClausula(Aux, pobjstcPed, lMensagem) Then GoTo ControleErro
        End If
        lLocal = "11 " & lI
        lLocal = "12 " & lI
        lrst.MoveNext
    Loop

    mfVisualizarClausulas = True
    Exit Function
ControleErro:

    If gM46V999.gfPreenchido(lMensagem) = True Then
        Call gM46V999.gpGraLog(1, "mfVisualizarClausulas" & lLocal & Chr(13) & lMensagem)
    Else
        Call gM46V999.gpGraLog(1, "mfVisualizarClausulas" & lLocal & Chr(13) & Err.Number & " - " & Err.Description)
    End If
    On Error Resume Next

    On Error GoTo 0
End Function
Private Function mfImprimirClausula(pArquivo As String, pobjstc As stcA46V702B, ByRef pMensagem As String)
    Dim lobjWord As Word.Application
    Dim lAux As String

    On Error GoTo ControleErro
    lAux = 1
    Set lobjWord = New Word.Application
    lAux = 2
    lobjWord.Documents.Open (pArquivo)
    lAux = 3
    If Not mfImprimir(lobjWord, pobjstc, pMensagem, True) Then GoTo ControleErro
    lAux = 4
    lobjWord.Quit wdDoNotSaveChanges
    lAux = 5
    mfImprimirClausula = True
    Exit Function
ControleErro:
    pMensagem = "mfImprimirClausula " & lAux & " " & Err.Description
End Function
Private Function mfJuntarArquivos(pobjWord As Word.Application, pArquivo As String, _
                                  pInserirQuebra As Boolean, pMensagem As String, pUltimoArq As Boolean, _
                                  Optional pSessao, Optional pTopMargem, Optional pBottomMargem, _
                                  Optional pLeftMargem, Optional pRightMargem, _
                                  Optional pHeaderDistance, Optional pFooterDistance, _
                                  Optional pPageWidth, Optional pPageHeigh) As Boolean

    Dim lSessao As Integer

    On Error GoTo ControleErro

    If pInserirQuebra Then
        pobjWord.ActiveDocument.Range(pobjWord.ActiveDocument.Range.End - 1).InsertBreak wdSectionBreakNextPage
        If pobjWord.ActiveWindow.View.SplitSpecial <> wdPaneNone Then
            pobjWord.ActiveWindow.Panes(2).Close
        End If
        If pobjWord.ActiveWindow.ActivePane.View.Type = wdNormalView Or _
           pobjWord.ActiveWindow.ActivePane.View.Type = wdOutlineView Or _
           pobjWord.ActiveWindow.ActivePane.View.Type = wdMasterView Then
            pobjWord.ActiveWindow.ActivePane.View.Type = wdPageView
        End If
        pobjWord.ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
        pobjWord.ActiveDocument.Sections(pobjWord.ActiveDocument.Sections.Count).Footers(wdHeaderFooterPrimary).LinkToPrevious = _
        Not pobjWord.ActiveDocument.Sections(pobjWord.ActiveDocument.Sections.Count - 1).Footers(wdHeaderFooterPrimary).LinkToPrevious
        pobjWord.ActiveDocument.Sections(pobjWord.ActiveDocument.Sections.Count).Footers(wdHeaderFooterPrimary).LinkToPrevious = _
        Not pobjWord.ActiveDocument.Sections(pobjWord.ActiveDocument.Sections.Count - 1).Footers(wdHeaderFooterPrimary).LinkToPrevious
        pobjWord.ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
    End If

    pobjWord.ActiveDocument.Range(pobjWord.ActiveDocument.Range.End - 1).InsertFile filename:=pArquivo, Range:="", _
                                                                                    ConfirmConversions:=False, Link:=False, Attachment:=False

    If IsMissing(pSessao) Then
        lSessao = pobjWord.ActiveDocument.Sections.Count
    Else
        lSessao = pSessao
    End If

    If IsMissing(pTopMargem) = False Then
        pobjWord.ActiveDocument.Sections(lSessao).PageSetup.TopMargin = pobjWord.CentimetersToPoints(pTopMargem)
    End If
    If IsMissing(pBottomMargem) = False Then
        pobjWord.ActiveDocument.Sections(lSessao).PageSetup.BottomMargin = pobjWord.CentimetersToPoints(pBottomMargem)
    End If
    If IsMissing(pLeftMargem) = False Then
        pobjWord.ActiveDocument.Sections(lSessao).PageSetup.LeftMargin = pobjWord.CentimetersToPoints(pLeftMargem)
    End If
    If IsMissing(pRightMargem) = False Then
        pobjWord.ActiveDocument.Sections(lSessao).PageSetup.RightMargin = pobjWord.CentimetersToPoints(pRightMargem)
    End If
    If IsMissing(pHeaderDistance) = False Then
        pobjWord.ActiveDocument.Sections(lSessao).PageSetup.HeaderDistance = pobjWord.CentimetersToPoints(pHeaderDistance)
    End If
    If IsMissing(pFooterDistance) = False Then
        pobjWord.ActiveDocument.Sections(lSessao).PageSetup.FooterDistance = pobjWord.CentimetersToPoints(pFooterDistance)
    End If
    pobjWord.ActiveDocument.Sections(lSessao).PageSetup.PaperSize = wdPaperCustom
    If IsMissing(pPageWidth) = False Then
        pobjWord.ActiveDocument.Sections(lSessao).PageSetup.PageWidth = pobjWord.CentimetersToPoints(pPageWidth)
    End If
    If IsMissing(pPageHeigh) = False Then
        pobjWord.ActiveDocument.Sections(lSessao).PageSetup.PageHeight = pobjWord.CentimetersToPoints(pPageHeigh)
    End If

    pobjWord.ActiveDocument.PageSetup.PaperSize = wdPaperCustom


    mfJuntarArquivos = True
    Exit Function
ControleErro:
    pMensagem = Err.Number & " - " & Err.Description
End Function
Private Sub Form_Unload(Cancel As Integer)
    mTipoDocumento = 0
End Sub
Public Function mfGerarFichaCorretor(pobjstcPed As stcA46V702B) As Boolean
    Dim lNomeArq As String
    Dim lFileObject As Scripting.FileSystemObject
    Dim lArquivo As Scripting.TextStream
    Dim lRegistro As String
    Dim lobjWord As Word.Application
    Dim lNomArqFicha As String
    Dim lMensagem As String
    Dim lArq As String
    Dim objDoc As Word.Document
    Dim lAuxiliar As String
    Dim lCount As Long


    On Error GoTo ControleErro


    lNomeArq = mLetraMapeamento & ":\docdiversos\" & Format(pobjstcPed.objstcCtrlEmis.Num_Apol, "0000000000") & Format(pobjstcPed.objstcCtrlEmis.Num_Endo, "000000") & "FSCD.txt"
    If Dir(lNomeArq) = "" Then
        lMensagem = "A Ficha do corretor da apólice " & _
                    Format(pobjstcPed.objstcCtrlEmis.Num_Apol, "0000000000") & "/" & _
                    Format(pobjstcPed.objstcCtrlEmis.Num_Endo, "000000") & _
                  " não foi localizada." & vbCrLf & vbCrLf & vbCrLf & _
                    "Sr. operador, abrir uma ficha para que seja analisado o motivo."
        Call gM46V999.gpGraLog(0, lMensagem)
        Call MsgBox(lMensagem, vbCritical)
        mfGerarFichaCorretor = True
        Exit Function
    End If
    lNomeArq = Dir(lNomeArq)
    lNomeArq = mLetraMapeamento & ":\docdiversos\" & lNomeArq
    Set lFileObject = New FileSystemObject
    Set lArquivo = lFileObject.OpenTextFile(lNomeArq)

    'Lê a primeira linha (sem informação)
    lRegistro = lArquivo.ReadLine
    Set lobjWord = New Word.Application

    'If gfForNormalDot(lobjWord, lMensagem) = False Then
    '    GoTo ControleErro
    'End If

    lobjWord.Visible = False
    lobjWord.Application.DisplayAlerts = wdAlertsNone

    lArq = "Ficha_Corretor.doc"
    If Dir(gM46V999.gPathArquivos & lArq) = "" Then
        Call gM46V999.gpGraLog(1, "Arquivo da ficha do corretor não foi localizado (" & gM46V999.gPathArquivos & lArq & ")")
        Exit Function
    End If

    Call FileCopy(gM46V999.gPathArquivos & lArq, gM46V999.gAppPath & lArq)

    Set objDoc = lobjWord.Documents.Open(gM46V999.gAppPath & lArq)
    lNomArqFicha = Format(pobjstcPed.objstcCtrlEmis.Num_Apol, "0000000000") & Format(pobjstcPed.objstcCtrlEmis.Num_Endo, "000000") & "_FSC.doc"

    'Lê próxima linha (Primeiro corretor)
    lRegistro = lArquivo.ReadLine
    lAuxiliar = Trim(Mid(lRegistro, 67, 25))
    Call mpReplaceWord(objDoc, "@@Corretor1@@", lAuxiliar)
    lAuxiliar = Trim(Mid(lRegistro, 92, 10))
    Call mpReplaceWord(objDoc, "@@Com1@@", lAuxiliar)

    'Lê próxima linha (Segundo corretor)
    lRegistro = lArquivo.ReadLine
    lAuxiliar = Trim(Mid(lRegistro, 67, 25))
    Call mpReplaceWord(objDoc, "@@Corretor2@@", lAuxiliar)
    lAuxiliar = Trim(Mid(lRegistro, 92, 10))
    Call mpReplaceWord(objDoc, "@@Com2@@", lAuxiliar)

    'Lê próxima linha (Terceiro corretor)
    lRegistro = lArquivo.ReadLine
    lAuxiliar = Trim(Mid(lRegistro, 67, 25))
    Call mpReplaceWord(objDoc, "@@Corretor3@@", lAuxiliar)
    lAuxiliar = Trim(Mid(lRegistro, 92, 10))
    Call mpReplaceWord(objDoc, "@@Com3@@", lAuxiliar)

    'Lê próxima linha
    lRegistro = lArquivo.ReadLine

    'Lê próxima linha
    lRegistro = lArquivo.ReadLine

    'Lê próxima linha (Registro com apólice-vigência, ramo...)
    lRegistro = lArquivo.ReadLine
    lAuxiliar = Trim(Mid(lRegistro, 1, 13))
    Call mpReplaceWord(objDoc, "@@Emissao@@", lAuxiliar)
    lAuxiliar = Trim(Mid(lRegistro, 13, 16))
    Call mpReplaceWord(objDoc, "@@Apolice@@", lAuxiliar)
    lAuxiliar = Trim(Mid(lRegistro, 30, 6))
    Call mpReplaceWord(objDoc, "@@Compl@@", lAuxiliar)
    lAuxiliar = pobjstcPed.gDesc_Ramo

    Call mpReplaceWord(objDoc, "@@Ramo@@", lAuxiliar)
    lAuxiliar = Trim(Mid(lRegistro, 56, 14))
    Call mpReplaceWord(objDoc, "@@IniVig@@", lAuxiliar)
    lAuxiliar = Trim(Mid(lRegistro, 70, 14))
    Call mpReplaceWord(objDoc, "@@TerVig@@", lAuxiliar)
    lAuxiliar = Trim(Mid(lRegistro, 84, 9))
    Call mpReplaceWord(objDoc, "@@Orgao@@", lAuxiliar)
    lAuxiliar = Trim(Mid(lRegistro, 90, 9))
    Call mpReplaceWord(objDoc, "@@Produtor@@", lAuxiliar)

    'Lê próxima linha (vazia)
    lRegistro = lArquivo.ReadLine

    'Lê próxima linha (segurado)
    lRegistro = lArquivo.ReadLine
    lAuxiliar = Trim(Mid(lRegistro, 1, 68))
    Call mpReplaceWord(objDoc, "@@Segurado@@", lAuxiliar)
    lAuxiliar = Trim(Mid(lRegistro, 69, 20))
    Call mpReplaceWord(objDoc, "@@CNPJ@@", lAuxiliar)

    'Lê próxima linha (vazia)
    lRegistro = lArquivo.ReadLine

    'Lê próxima linha (endereço)
    lRegistro = lArquivo.ReadLine
    lAuxiliar = Trim(Mid(lRegistro, 1, 200))
    Call mpReplaceWord(objDoc, "@@Endereco@@", lAuxiliar)

    'Lê próxima linha (continuação endereço)
    lRegistro = lArquivo.ReadLine
    lAuxiliar = Trim(Mid(lRegistro, 1, 200))
    Call mpReplaceWord(objDoc, "@@ContEndereco@@", lAuxiliar)

    'Lê próxima linha (vazia)
    lRegistro = lArquivo.ReadLine

    For lCount = 1 To 12
        'Lê próxima linha (vazia)
        lRegistro = lArquivo.ReadLine
        'Lê próxima linha (vazia)
        lRegistro = lArquivo.ReadLine
        'Lê próxima linha (linha da parcela)
        lRegistro = lArquivo.ReadLine
        lAuxiliar = Trim(Mid(lRegistro, 1, 10))
        Call mpReplaceWord(objDoc, "@@Parc" & lCount & "@@", lAuxiliar)
        lAuxiliar = Trim(Mid(lRegistro, 11, 10))
        Call mpReplaceWord(objDoc, "@@Venc" & lCount & "@@", lAuxiliar)
        lAuxiliar = Trim(Mid(lRegistro, 21, 26))
        Call mpReplaceWord(objDoc, "@@Premio" & lCount & "@@", lAuxiliar)
        lAuxiliar = Trim(Mid(lRegistro, 47, 11))
        Call mpReplaceWord(objDoc, "@@Cobert" & lCount & "@@", lAuxiliar)
        lAuxiliar = Trim(Mid(lRegistro, 58, 7))
        Call mpReplaceWord(objDoc, "@@Perc" & lCount & "@@", lAuxiliar)
        lAuxiliar = Trim(Mid(lRegistro, 65, 12))
        Call mpReplaceWord(objDoc, "@@ValCom" & lCount & "@@", lAuxiliar)
        lAuxiliar = Trim(Mid(lRegistro, 77, 8))
        Call mpReplaceWord(objDoc, "@@CodSeg" & lCount & "@@", lAuxiliar)
        lAuxiliar = Trim(Mid(lRegistro, 85, 13))
        Call mpReplaceWord(objDoc, "@@PercSeg" & lCount & "@@", lAuxiliar)
    Next lCount

    ''Lê próxima linha (Valor total)
    lRegistro = lArquivo.ReadLine
    lRegistro = lArquivo.ReadLine
    lRegistro = lArquivo.ReadLine
    lAuxiliar = Trim(Mid(lRegistro, 1, 46))
    Call mpReplaceWord(objDoc, "@@TotalPremio@@", lAuxiliar)
    lAuxiliar = Trim(Mid(lRegistro, 47, 52))
    Call mpReplaceWord(objDoc, "@@TotalPer@@", lAuxiliar)

    lobjWord.ActiveDocument.SaveAs gM46V999.gAppPath & lNomArqFicha
    If Not mfImprimir(lobjWord, pobjstcPed, lMensagem, False) Then GoTo ControleErro

    'Finalizar
    lobjWord.ActiveDocument.Close
    lobjWord.Quit
    Set lobjWord = Nothing
    Set objDoc = Nothing

    On Error Resume Next
    Kill gM46V999.gAppPath & lNomArqFicha
    mfGerarFichaCorretor = True
    Exit Function
ControleErro:
    If Not gM46V999.gfPreenchido(lMensagem) Then
        lMensagem = Err.Description
    End If
    Call gM46V999.gpGraLog(1, lMensagem)
End Function


