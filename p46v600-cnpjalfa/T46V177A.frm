VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmT46V177A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Conferência da proposta"
   ClientHeight    =   9480
   ClientLeft      =   360
   ClientTop       =   945
   ClientWidth     =   9105
   Icon            =   "T46V177A.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9480
   ScaleWidth      =   9105
   ShowInTaskbar   =   0   'False
   Begin RichTextLib.RichTextBox rtbTexto 
      Height          =   8895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   15690
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"T46V177A.frx":000C
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
   Begin Threed.SSCommand cmdOK 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   345
      Left            =   4545
      TabIndex        =   1
      Top             =   9030
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   617
      _StockProps     =   78
      Caption         =   "&OK"
      ForeColor       =   8388608
   End
   Begin Threed.SSCommand cmdVoltar 
      Height          =   345
      Left            =   3450
      TabIndex        =   2
      Top             =   9030
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   617
      _StockProps     =   78
      Caption         =   "&Voltar"
      ForeColor       =   8388608
   End
End
Attribute VB_Name = "frmT46V177A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objPedido As clsA46V702A    'Objeto da classe clsA46V702A (TAB_PED)
Private objstcPedido As stcA46V702B    'Objeto da classe de estrutura stcA46V702B (TAB_PED)

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
Private Sub cmdOk_Click()
    frmT46V101A.pDadosOK = True
    Unload Me
End Sub

Private Sub cmdVoltar_Click()
    frmT46V101A.pDadosOK = False
    Unload Me
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

    Dim lobjItem As clsA46V708A    'Objeto da classe clsA46V708A (TAB_PED_LOC)
    Dim lobjstcItem As stcA46V708B    'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
    Dim lobjCobertura As clsA46V704A    'Objeto da classe clsA46V704A (TAB_PED_COBERT)
    Dim lobjstcCobertura As stcA46V704B    'Objeto da classe de estrutura stcA46V704B (TAB_PED_COBERT)
    Dim lobjEndereco As clsA46V707A    'Objeto da classe clsA46V707A (TAB_PED_ENDER)
    Dim lobjstcEndereco As stcA46V707B    'Objeto da classe de estrutura stcA46V707B (TAB_PED_ENDER)

    Screen.MousePointer = vbHourglass
    On Error GoTo ControleErro
    '1. Centraliza form.
    Call gM46V999.gpCenForm(Me)

    '2. Inicia processo.
    mQuaLinhas = 0
    Call mpPosLinha(String(80, "="))
    Call mpPosLinha("Dados da Proposta - SYAS Multiseguros.")
    Call mpPosLinha(String(80, "="))
    Call mpPosLinha(" ")

    '3. Obtém dados da proposta.
    Set objstcPedido = frmT46V101A.objstcPedido

    '4. Exibe dados da proposta.
    lPriProposta = True

    '***==== DADOS DA PROPOSTA
    If lNosNumero <> objstcPedido.NOSSO_NUMERO Then
        lNosNumero = objstcPedido.NOSSO_NUMERO
        If gM46V999.gfPreenchido(objstcPedido.NOSSO_NUMERO) Then
            Call mpPosLinha("Proposta ............." & objstcPedido.NOSSO_NUMERO)
        Else
            Call mpPosLinha("Proposta .............não informada")
        End If
        If gM46V999.gfPreenchido(objstcPedido.Num_Ped) Then
            Call mpPosLinha("Pedido................" & objstcPedido.Num_Ped)
        Else
            Call mpPosLinha("Pedido................não informado")
        End If
        If gM46V999.gfPreenchido(objstcPedido.Cod_Ramo) Then
            Call mpPosLinha("Ramo.................." & objstcPedido.Cod_Ramo)
        Else
            Call mpPosLinha("Ramo..................não informado")
        End If
        If gM46V999.gfPreenchido(objstcPedido.Nom_Segurado) Then
            Call mpPosLinha("Segurado.............." & Trim$(objstcPedido.Nom_Segurado))
        Else
            Call mpPosLinha("Segurado..............não informado.")
        End If
        If gM46V999.gfPreenchido(objstcPedido.gCNPJ_CPF_Formatado) Then
            Call mpPosLinha("CPF/CNPJ ............." & Trim$(objstcPedido.gCNPJ_CPF_Formatado))
        Else
            Call mpPosLinha("CPF/CNPJ .............não informado.")
        End If
        If gM46V999.gfPreenchido(objstcPedido.Dat_Ini_Vig) Then
            Call mpPosLinha("Início de vigência...." & gM46V999.gfForData(2, objstcPedido.Dat_Ini_Vig))
        Else
            Call mpPosLinha("Início de vigência....não informado")
        End If
        If gM46V999.gfPreenchido(objstcPedido.Dat_Ter_Vig) Then
            Call mpPosLinha("Término de vigência..." & gM46V999.gfForData(2, objstcPedido.Dat_Ter_Vig))
        Else
            Call mpPosLinha("Término de vigência...não informado")
        End If


        '5. Obtém dados de endereço do Segurado.
        If gM46V999.gfPreenchido(objstcPedido.objstcPedEndereco.Cod_Tip_Logr) Then
            Call mpPosLinha("Endereço.............." & objstcPedido.objstcPedEndereco.Cod_Tip_Logr & _
                          " " & objstcPedido.objstcPedEndereco.Nom_Logr & _
                            ", " & objstcPedido.objstcPedEndereco.Dsc_Compl_Logr)
        Else
            Call mpPosLinha("Endereço..............não informado")
        End If


        '6. Exibe valor dos prêmios
        If gM46V999.gfPreenchido(objstcPedido.VAL_PR) Then
            Call mpPosLinha("Prêmio Líquido........" & Format(objstcPedido.VAL_PR, "0.00"))

        Else
            Call mpPosLinha("Prêmio Líquido........não informado")
        End If
        If gM46V999.gfPreenchido(objstcPedido.Val_Cust_Emis) Then
            Call mpPosLinha("Custo ................" & Format(objstcPedido.Val_Cust_Emis, "0.00"))
        Else
            Call mpPosLinha("Custo ................não informado")
        End If
        If gM46V999.gfPreenchido(objstcPedido.gValorTotal) Then
            Call mpPosLinha("Prêmio Total ........." & Format(objstcPedido.gValorTotal, "0.00"))
        Else
            Call mpPosLinha("Prêmio Total .........não informado")
        End If

        '7. Exibe dados de pagamento
        If gM46V999.gfPreenchido(objstcPedido.objstcCtrlEmis.Cod_Iden_Pag_Antec) Then
            If objstcPedido.objstcCtrlEmis.Cod_Iden_Pag_Antec = e_SimNao_Sim Then
                Call mpPosLinha("Forma de Pagamento... Antecipado - " & IIf(objstcPedido.Num_Cta_Bas_Cor <> 0, "Débito em Conta", "Carnê"))
                If objstcPedido.objstcCtrlEmis.Num_Bloq <> 0 Then Call mpPosLinha("                     Bloqueto:" & objstcPedido.objstcCtrlEmis.Num_Bloq)
                If objstcPedido.Num_Cta_Bas_Cor <> 0 Then mpPosLinha ("                      Banco:" & objstcPedido.Cod_Bco_Cobr & " Agência:" & objstcPedido.Cod_Agen_Bas_Cor & " Conta:" & objstcPedido.Num_Cta_Bas_Cor)
            Else
                Call mpPosLinha("Forma de Pagamento... A 30 dias  - " & IIf(objstcPedido.Num_Cta_Bas_Cor <> 0, "Débito em Conta", "Carnê"))
                If objstcPedido.objstcCtrlEmis.Num_Bloq <> 0 Then Call mpPosLinha("                     Bloqueto:" & objstcPedido.objstcCtrlEmis.Num_Bloq)
                If objstcPedido.Num_Cta_Bas_Cor <> 0 Then mpPosLinha ("                      Banco:" & objstcPedido.Cod_Bco_Cobr & " Agência:" & objstcPedido.Cod_Agen_Bas_Cor & " Conta:" & objstcPedido.Num_Cta_Bas_Cor)
            End If
        Else
            Call mpPosLinha("Forma de Pagamento....não informado")
        End If


        '8. Obtém dados dos itens.
        Set lobjItem = New clsA46V708A
        Set lobjItem.M46V999 = gM46V999

        For Each lobjstcItem In objstcPedido.ColecaoPedLoc
            Call mpPosLinha(" ")
            Call mpPosLinha(String(80, "-"))
            Call mpPosLinha("Dados do Item")
            Call mpPosLinha(String(80, "-"))
            Call mpPosLinha(" ")

            If gM46V999.gfPreenchido(lobjstcItem.Num_Item) Then
                Call mpPosLinha("Item.................." & lobjstcItem.Num_Item)
            Else
                Call mpPosLinha("Item..................não informado")
            End If
            Call mpPosLinha("Local de Risco........" & lobjstcItem.Cod_Tip_Logr_Risco & _
                          " " & lobjstcItem.Nom_Logr_Risco & _
                            ", " & lobjstcItem.objstcPedInsp.Num_Logr_Risco)

            If gM46V999.gfPreenchido(lobjstcItem.Dsc_Compl_Risco) Then
                Call mpPosLinha("Complemento..........." & lobjstcItem.Dsc_Compl_Risco)
            Else
                Call mpPosLinha("Complemento...........não informado")
            End If
            If gM46V999.gfPreenchido(lobjstcItem.Num_Cep_Risco) Then
                Call mpPosLinha("CEP..................." & lobjstcItem.gCEP_Formatado)
            Else
                Call mpPosLinha("CEP...................não informado")
            End If
            If gM46V999.gfPreenchido(lobjstcItem.Nom_Cid_Risco) Then
                Call mpPosLinha("Cidade................" & lobjstcItem.Nom_Cid_Risco)
            Else
                Call mpPosLinha("Cidade................não informada")
            End If
            If gM46V999.gfPreenchido(lobjstcItem.Sig_UF_Risco) Then
                Call mpPosLinha("UF...................." & lobjstcItem.Sig_UF_Risco)
            Else
                Call mpPosLinha("UF....................não informada")
            End If

            If gM46V999.gfPreenchido(lobjstcItem.Cod_Clasf) And gM46V999.gfPreenchido(lobjstcItem.Cod_Compl_Clasf) Then
                Call mpPosLinha("Classificação........." & lobjstcItem.Cod_Clasf & "-" & lobjstcItem.Cod_Compl_Clasf & " " & lobjstcItem.gNom_Class)
            Else
                Call mpPosLinha("Classificação.........não informada")
            End If

            '9. Obtém dados das coberturas.
            Set lobjCobertura = New clsA46V704A
            Set lobjCobertura.M46V999 = gM46V999

            Call mpPosLinha(" ")
            Call mpPosLinha("Dados da(s) Cobertura(s)")
            Call mpPosLinha(String(80, "-"))
            Call mpPosLinha(" ")

            For Each lobjstcCobertura In lobjstcItem.ColecaoPedCobert
                If gM46V999.gfPreenchido(lobjstcCobertura.Cod_Cobert) Then
                    Call mpPosLinha("Cobertura..." & Space(3 - Len(Mid(lobjstcCobertura.Cod_Cobert, 1, 3))) & lobjstcCobertura.Cod_Cobert & "-" & Mid(lobjstcCobertura.gNomeCobertura, 1, 40) & Space(40 - Len(Mid(lobjstcCobertura.gNomeCobertura, 1, 40))) & "      IS..." & Format(lobjstcCobertura.Val_IS, "0.00"))
                Else
                    Call mpPosLinha("Cobertura...não informada")
                End If
            Next lobjstcCobertura
        Next lobjstcItem
    End If

    Call mpPosLinha(String(80, "="))
    Call mpPosLinha(" ")
    Call mpPosLinha("FIM " & String(76, "="))
    Set lobjstcCobertura = Nothing
    Set lobjCobertura = Nothing
    Set lobjstcItem = Nothing
    Set lobjItem = Nothing
    Set lobjEndereco = Nothing
    Set lobjstcEndereco = Nothing
    Set objstcPedido = Nothing
    Set objPedido = Nothing

    Screen.MousePointer = vbDefault
    Exit Sub
ControleErro:
    Call gM46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gM46V999.gpGraLog(0, "T46V1766A.Load ERRO - " & Err & " " & Err.Description)
End Sub
