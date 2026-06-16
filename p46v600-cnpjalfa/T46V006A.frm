VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmT46V006A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Protocolo de transmissão"
   ClientHeight    =   5625
   ClientLeft      =   360
   ClientTop       =   945
   ClientWidth     =   9015
   Icon            =   "T46V006A.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog cdgImpressoras 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      PrinterDefault  =   0   'False
   End
   Begin RichTextLib.RichTextBox rtbTexto 
      Height          =   5100
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   8996
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"T46V006A.frx":000C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCommand cmdSair 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   345
      Left            =   4550
      TabIndex        =   2
      Top             =   5200
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   617
      _StockProps     =   78
      Caption         =   "&Sair"
      ForeColor       =   8388608
   End
   Begin Threed.SSCommand cmdImprimir 
      Height          =   350
      Left            =   3450
      TabIndex        =   1
      Top             =   5200
      Width           =   1000
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   617
      _StockProps     =   78
      Caption         =   "&Imprimir"
      ForeColor       =   8388608
   End
End
Attribute VB_Name = "frmT46V006A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mQuaLinhas As Long    'Quantidade de linhas.
Private Sub mpPosLinha(ByVal pMensagem As String)
'Procedure: posiciona linha de mensagem.

    Dim lAuxiliar As String    'Auxiliar.

    With rtbTexto
        lAuxiliar = pMensagem
        If Len(lAuxiliar) < 81 Then
            If gM46V999.gfPreenchido(.Text) Then
                .Text = .Text & vbLf & lAuxiliar
            Else
                .Text = lAuxiliar
            End If
            mQuaLinhas = mQuaLinhas + 1
        Else
            If gM46V999.gfPreenchido(.Text) Then
                .Text = .Text & vbLf & Left$(lAuxiliar, 77) & "..."
            Else
                .Text = Left$(lAuxiliar, 77) & "..."
            End If
            mQuaLinhas = mQuaLinhas + 1
            lAuxiliar = Mid$(lAuxiliar, 78)
            Do
                If Len(lAuxiliar) < 77 Then
                    .Text = .Text & vbLf & "..." & lAuxiliar
                    mQuaLinhas = mQuaLinhas + 1
                    Exit Do
                End If
                .Text = .Text & vbLf & "..." & Left$(lAuxiliar, 74) & "..."
                mQuaLinhas = mQuaLinhas + 1
                lAuxiliar = Mid$(lAuxiliar, 5)
            Loop
            .Text = .Text & vbLf & " "
            mQuaLinhas = mQuaLinhas + 1
        End If
        .Refresh
    End With
End Sub
Private Sub cmdImprimir_Click()
    Dim lMensagem As String    'Mensagem.
    Dim lQuaPaginas As Integer  'Quantidade de páginas.

'1. Confirma com usuário.
    If mQuaLinhas > 85 Then
        lQuaPaginas = CInt(mQuaLinhas / 85)
        Select Case lQuaPaginas
        Case Is <= 5
            lMensagem = "AGE0013 - A impressão do protocolo resultará em mais de uma página. " & _
                        vbLf & vbLf & "Confirma impressão do protocolo?"
        Case Else
            lMensagem = "AGE0013 - A impressão do protocolo resultará em mais de uma página " & _
                        "(aproximadamente " & lQuaPaginas & " páginas)." & vbLf & vbLf & _
                        "Confirma impressão do protocolo?"
        End Select
        If MsgBox(lMensagem, vbCritical + vbYesNo + vbDefaultButton2) = vbYes Then
            Call gM46V999.gpGraLog(0, lMensagem & " SIM.")
        Else
            Call gM46V999.gpGraLog(0, lMensagem & " NÃO.")
            Exit Sub
        End If
    End If

    '2. Seleciona impressora e imprime protocolo.
    With cdgImpressoras
        .CancelError = True
        .PrinterDefault = False
        .Flags = cdlPDPrintSetup + cdlPDReturnDC
        On Error Resume Next
        .ShowPrinter
        If Err <> 0 Then    'O usuário pressionou o botão Cancel.
            On Error GoTo 0
            Exit Sub
        End If
        On Error GoTo 0
        rtbTexto.SelPrint .hDC
    End With

    '3. Encerra processo.
    Call gM46V999.gpGraLog(1, "AGE0002 - Processo concluído: impressão.")
    Unload Me
End Sub
Private Sub cmdSair_Click()
    Unload Me
End Sub
Private Sub Form_Activate()
    frmT46V001A.panForm = Mid$(Me.Name, 4)
End Sub
Private Sub Form_Load()
    Dim lDir As String    'Utilizado com a função DIR.
    Dim lMensagem As String     'Mensagem.
    Dim lNosNumero As String        'Nosso numero
    Dim lNumArquivo As Integer      'Utilizado com FreeFile.
    Dim lPriProposta As Boolean       'Primeira proposta.
    Dim lRegistro As String     'Registro.
    Dim lobjTransm As clsA46V718A    'Objeto da classe clsA46V718A (TAB_PED_TRANSM)
    Dim lobjstcTransm As stcA46V718B    'Objeto da classe de estrutura stcA46V718B (TAB_PED_TRANSM)
    Dim lCount As Long    'Contador.
    Dim clsA99V001 As clsA99V001
    Dim lDigito As String
    Screen.MousePointer = vbHourglass
    On Error GoTo ControleErro
    '1. Centraliza form.
    Call gM46V999.gpCenForm(Me)

    '2. Inicia processo.
    mQuaLinhas = 0
    Call mpPosLinha("YASUDA - PROTOCOLO DE TRANSMISSÃO - " & Format$(Now, "dd/mm/yyyy hh:nn:ss") & ".")
    Call mpPosLinha(" ")

    '4. Obtém registros de protocolo referente ao nosso número.
    Set lobjTransm = New clsA46V718A
    Set lobjTransm.M46V999 = gM46V999

    Set lobjstcTransm = New stcA46V718B
    lobjstcTransm.NOSSO_NUMERO = gM46V111.gComNosNumero

    If Not lobjTransm.mfCarregarColecao(lobjstcTransm, "", "") Then
        Call gM46V999.gpGraLog(2, lobjTransm.UltimoErro)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If lobjTransm.Colecao.Count = 0 Then
        Call gM46V999.gpGraLog(2, "FNL0014 - Registro de protocolo de transmissão não localizado (pedido " & _
                                  gM46V111.gNumPedido & " - nosso número " & gM46V111.gComNosNumero & ").")
        Set lobjTransm = Nothing
        Set lobjstcTransm = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    'c) Obtém dados da proposta transmitida (detalhes - erro).
    '*************************Ver este ponto

    '5. Processa registro de protocolo.
    lPriProposta = True
    Call mpPosLinha("Protocolo de transmissão - SYAS Multiseguros.")
    Call mpPosLinha(" ")
    For Each lobjstcTransm In lobjTransm.Colecao
        'a) Exibe dados
        'Call mpPosLinha("Transmissão da Yasuda: " & lobjstcTransm.Nom_Arq_Retorno & ".")
        If lNosNumero <> lobjstcTransm.NOSSO_NUMERO Then
            lNosNumero = lobjstcTransm.NOSSO_NUMERO
            If gM46V999.gfPreenchido(lobjstcTransm.NOSSO_NUMERO) Then
                Call mpPosLinha("Proposta recebida....." & lobjstcTransm.NOSSO_NUMERO)
            Else
                Call mpPosLinha("Proposta recebida.....não informado")
            End If
            If gM46V999.gfPreenchido(lobjstcTransm.Num_Ped) Then
                Call mpPosLinha("Pedido................" & lobjstcTransm.Num_Ped)
            Else
                Call mpPosLinha("Pedido................não informado")
            End If
            If gM46V999.gfPreenchido(lobjstcTransm.Cod_Ramo) Then
                Call mpPosLinha("Ramo.................." & lobjstcTransm.Cod_Ramo)
            Else
                Call mpPosLinha("Ramo..................não informado")
            End If
            If gM46V999.gfPreenchido(lobjstcTransm.Nom_Segurado) Then
                Call mpPosLinha("Segurado.............." & Trim$(lobjstcTransm.Nom_Segurado))
            Else
                Call mpPosLinha("Segurado..............não informado.")
            End If
            If gM46V999.gfPreenchido(lobjstcTransm.Dat_Ini_Vig) Then
                Call mpPosLinha("Início de vigência...." & gM46V999.gfForData(2, lobjstcTransm.Dat_Ini_Vig))
            Else
                Call mpPosLinha("Início de vigência....não informado")
            End If
            If gM46V999.gfPreenchido(lobjstcTransm.Dat_Ter_Vig) Then
                Call mpPosLinha("Término de vigência..." & gM46V999.gfForData(2, lobjstcTransm.Dat_Ter_Vig))
            Else
                Call mpPosLinha("Término de vigência...não informado")
            End If
            If gM46V999.gfPreenchido(lobjstcTransm.Nom_Arq_Envio) Then
                Call mpPosLinha("Arquivo enviado......." & lobjstcTransm.Nom_Arq_Envio)
            Else
                Call mpPosLinha("Arquivo enviado.......não informado")
            End If
            If gM46V999.gfPreenchido(lobjstcTransm.Dat_Hor_Envio) Then
                Call mpPosLinha("Data/Hora envio......." & gM46V999.gfForData(4, lobjstcTransm.Dat_Hor_Envio) & " horas.")
            Else
                Call mpPosLinha("Data/Hora envio.......não informado")
            End If
            If gM46V999.gfPreenchido(lobjstcTransm.Nom_Arq_Retorno) Then
                Call mpPosLinha("Arquivo de retorno...." & lobjstcTransm.Nom_Arq_Retorno)
            Else
                Call mpPosLinha("Arquivo de retorno....não informado")
            End If
            If gM46V999.gfPreenchido(lobjstcTransm.Dat_Hor_Retorno) Then
                Call mpPosLinha("Data/Hora retorno....." & gM46V999.gfForData(4, lobjstcTransm.Dat_Hor_Envio) & " horas.")
            Else
                Call mpPosLinha("Data/Hora retorno.....não informado")
            End If
            If Val(lobjstcTransm.Num_PI) <> 0 Then
                Set clsA99V001 = New clsA99V001
                Set clsA99V001.M46V999 = gM46V999

                If clsA99V001.DIG_YASUDA(lobjstcTransm.Num_PI, lDigito, lMensagem) = False Then
                    Exit Sub
                End If
                Set clsA99V001 = Nothing
                If lDigito = "A" Then lDigito = "0"
                Call mpPosLinha("Protocolo Interno....." & lobjstcTransm.Num_PI & "-" & lDigito)
                mpPosLinha (" ")
                mpPosLinha ("Os anexos, quando necessários, poderão ser enviados das seguintes formas:")
                mpPosLinha ("1) VIA FAX: Através  do  número  (0xx11) 3305-5664,  informando  o  número  do ")
                mpPosLinha ("protocolo interno " & lobjstcTransm.Num_PI & lDigito & " no atendimento eletrônico ou;")
                mpPosLinha ("2) VIA YasudaNet, 'ÁREA DO CORRETOR': Carregar o(s) documento(s) digitalizado(s)")
                mpPosLinha ("através  dos  links  Seguro  Demais Ramos  -->  Emissão.  Filtrar pelo Número do")
                mpPosLinha ("Protocolo (PI).")
                mpPosLinha (" ")
            End If
            Call mpPosLinha(" ")
            Select Case lobjstcTransm.Ind_Proposta
            Case "0"
                Call mpPosLinha("STATUS................Proposta não transmitida.")
            Case 1
                Call mpPosLinha("STATUS...Proposta eletrônica recebida pela Yasuda em " & _
                                gM46V999.gfForData(4, lobjstcTransm.Dat_Hor_Retorno) & " horas.")

            Case 2
                Call mpPosLinha("STATUS................Proposta apresentou erro.")
                Call mpPosLinha("É necessário retransmitir esta proposta.")
                Call mpPosLinha("Descrição do erro....." & lobjstcTransm.Tip_Emissao)  'Contém observação do erro
            Case Else
                Call mpPosLinha("STATUS................Proposta apresentou erro.")
                Call mpPosLinha("É necessário retransmitir esta proposta.")
                Call mpPosLinha("Descrição do erro.....Indicador de proposta desconhecido - comunicar Central de Atendimento.")
            End Select
            Call mpPosLinha(" ")
            Call mpPosLinha("Itens da proposta:")
        End If
        '***==== DADOS ITENS
        Call mpPosLinha(" ")
        If gM46V999.gfPreenchido(lobjstcTransm.Num_Item) Then
            Call mpPosLinha("Item.................." & lobjstcTransm.Num_Item)
        Else
            Call mpPosLinha("Item..................não informado")
        End If
        Call mpPosLinha("Logradouro............" & lobjstcTransm.Cod_Tip_Logr_Risco & _
                      " " & lobjstcTransm.Nom_Logr_Risco & _
                        ", " & lobjstcTransm.Num_Risco)

        If gM46V999.gfPreenchido(lobjstcTransm.Dsc_Compl_Risco) Then
            Call mpPosLinha("Complemento..........." & lobjstcTransm.Dsc_Compl_Risco)
        Else
            Call mpPosLinha("Complemento...........não informado")
        End If
        If gM46V999.gfPreenchido(lobjstcTransm.Num_Cep_Risco) Then
            Call mpPosLinha("CEP..................." & lobjstcTransm.Num_Cep_Risco)
        Else
            Call mpPosLinha("CEP...................não informado")
        End If
        If gM46V999.gfPreenchido(lobjstcTransm.Nom_Cid_Risco) Then
            Call mpPosLinha("Cidade................" & lobjstcTransm.Nom_Cid_Risco)
        Else
            Call mpPosLinha("Cidade................não informada")
        End If
        If gM46V999.gfPreenchido(lobjstcTransm.Sig_UF_Risco) Then
            Call mpPosLinha("UF...................." & lobjstcTransm.Sig_UF_Risco)
        Else
            Call mpPosLinha("UF....................não informada")
        End If
    Next lobjstcTransm
    Call mpPosLinha(String(80, "="))
    Call mpPosLinha(" ")
    Call mpPosLinha("FIM " & String(76, "="))
    Set lobjstcTransm = Nothing
    Set lobjTransm = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
ControleErro:
    Call gM46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gM46V999.gpGraLog(0, "T46V006A.Load ERRO - " & Err & " " & Err.Description)
End Sub
