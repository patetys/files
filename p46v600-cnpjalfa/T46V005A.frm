VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmT46V005A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Transmissão de propostas"
   ClientHeight    =   5625
   ClientLeft      =   2565
   ClientTop       =   2325
   ClientWidth     =   9000
   Icon            =   "T46V005A.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   Begin Threed.SSCommand cmdExcluir 
      Height          =   400
      Left            =   3700
      TabIndex        =   6
      Top             =   3000
      Width           =   1600
      _Version        =   65536
      _ExtentX        =   2822
      _ExtentY        =   706
      _StockProps     =   78
      Caption         =   "<- Excluir"
      ForeColor       =   8388608
   End
   Begin Threed.SSCommand cmdExcTodos 
      Height          =   400
      Left            =   3700
      TabIndex        =   5
      Top             =   2500
      Width           =   1600
      _Version        =   65536
      _ExtentX        =   2822
      _ExtentY        =   706
      _StockProps     =   78
      Caption         =   "<<- Excluir todos"
      ForeColor       =   8388608
   End
   Begin Threed.SSCommand cmdSelecionar 
      Height          =   400
      Left            =   3700
      TabIndex        =   4
      Top             =   2000
      Width           =   1600
      _Version        =   65536
      _ExtentX        =   2822
      _ExtentY        =   706
      _StockProps     =   78
      Caption         =   "Selecionar ->"
      ForeColor       =   8388608
   End
   Begin Threed.SSCommand cmdSelTodos 
      Height          =   400
      Left            =   3700
      TabIndex        =   3
      Top             =   1500
      Width           =   1600
      _Version        =   65536
      _ExtentX        =   2822
      _ExtentY        =   706
      _StockProps     =   78
      Caption         =   "Selecionar todos ->>"
      ForeColor       =   8388608
   End
   Begin Threed.SSCommand cmdOk 
      Height          =   405
      Left            =   3405
      TabIndex        =   10
      Top             =   5070
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   706
      _StockProps     =   78
      Caption         =   "&Ok"
      ForeColor       =   8388608
   End
   Begin Threed.SSCommand cmdCancelar 
      Cancel          =   -1  'True
      Height          =   405
      Left            =   4605
      TabIndex        =   11
      Top             =   5070
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   706
      _StockProps     =   78
      Caption         =   "&Cancelar"
      ForeColor       =   8388608
   End
   Begin Threed.SSFrame fraSelecao 
      Height          =   4650
      Left            =   5355
      TabIndex        =   7
      Top             =   315
      Width           =   3600
      _Version        =   65536
      _ExtentX        =   6350
      _ExtentY        =   8202
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin MSFlexGridLib.MSFlexGrid flxSelecao 
         Height          =   4005
         Left            =   105
         TabIndex        =   9
         Top             =   540
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   7064
         _Version        =   393216
         Rows            =   15
         Cols            =   7
         FixedCols       =   0
         ForeColor       =   8388608
         ForeColorFixed  =   8388608
         WordWrap        =   -1  'True
         AllowBigSelection=   0   'False
         FocusRect       =   0
         FillStyle       =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
      End
      Begin VB.Label lblX 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Propostas selecionadas para transmissão"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   105
         TabIndex        =   8
         Top             =   135
         Width           =   3405
      End
   End
   Begin Threed.SSFrame fraX 
      Height          =   4650
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   315
      Width           =   3600
      _Version        =   65536
      _ExtentX        =   6350
      _ExtentY        =   8202
      _StockProps     =   14
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin MSFlexGridLib.MSFlexGrid flxResultado 
         Height          =   4005
         Left            =   60
         TabIndex        =   2
         Top             =   540
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   7064
         _Version        =   393216
         Rows            =   15
         Cols            =   7
         FixedCols       =   0
         ForeColor       =   8388608
         ForeColorFixed  =   8388608
         WordWrap        =   -1  'True
         AllowBigSelection=   0   'False
         FocusRect       =   0
         FillStyle       =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
      End
      Begin VB.Label lblX 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Propostas calculadas"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   60
         TabIndex        =   1
         Top             =   135
         Width           =   3405
      End
   End
   Begin Threed.SSPanel panMensagem 
      Height          =   1005
      Left            =   -90
      TabIndex        =   14
      Top             =   5040
      Visible         =   0   'False
      Width           =   3405
      _Version        =   65536
      _ExtentX        =   5997
      _ExtentY        =   1764
      _StockProps     =   15
      ForeColor       =   12582912
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
   End
   Begin VB.Label lblUsuario 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Usuário"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   1260
      TabIndex        =   13
      Top             =   45
      Width           =   7695
   End
   Begin VB.Label lblX 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Usuário: "
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   0
      Left            =   135
      TabIndex        =   12
      Top             =   45
      Width           =   1095
   End
End
Attribute VB_Name = "frmT46V005A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private objPedido As clsA46V702A    'Objeto da classe clsA46V702A (TAB_PED)
Private objstcPedido As stcA46V702B   'Objeto da classe de estrutura stcA46V702B (TAB_PED)

Private Sub mpMonGrade(ByVal pGrade As MSFlexGrid)
'Procedure: monta grade.

    With pGrade
        'a) Posiciona quantidade de colunas e de linhas (apenas cabeçalho).
        .Cols = 7
        .Rows = 1
        'b) Posiciona linha de cabeçalho (com altura).
        .Row = 0
        .RowHeight(0) = 480
        'c) Posiciona cabeçalho nas colunas.
        .Col = 0
        .Text = "Pedido"
        .Col = 1
        .Text = "Segurado"
        .Col = 2
        .Text = "CNPJ/CPF"
        .Col = 3
        .Text = "Início de vigência"
        .Col = 4
        .Text = "Término de vigência"
        .Col = 5
        .Text = "Nosso número"
        .Col = 6
        .Text = "Código do corretor"
        'd) Posiciona largura das colunas.
        .ColWidth(0) = 1000
        .ColWidth(1) = 1500
        .ColWidth(2) = 1300
        .ColWidth(3) = 1000
        .ColWidth(4) = 1000
        .ColWidth(5) = 2000
        .ColWidth(6) = 0
        'e) Posiciona alinhamento das colunas.
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignCenterCenter
        .ColAlignment(4) = flexAlignCenterCenter
        .ColAlignment(5) = flexAlignLeftCenter
        .ColAlignment(6) = flexAlignLeftCenter
        .Visible = True
    End With
End Sub
Private Sub cmdCancelar_Click()
    Unload Me
End Sub
Private Sub cmdExcluir_Click()
    Dim lI As Byte        'Utilizado com For...Next.

    Screen.MousePointer = vbHourglass

    '1. Verifica se tem documento selecionado.
    If flxSelecao.Row = 0 Then
        Call gM46V999.gpGraLog(1, "FNP0014 - Documento não selecionado.")
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    '2. Coloca registro selecionado na grade de seleção.
    flxResultado.Rows = flxResultado.Rows + 1
    flxResultado.Row = flxResultado.Rows - 1
    For lI = 0 To 6
        flxSelecao.Col = lI
        flxResultado.Col = lI
        flxResultado.Text = flxSelecao.Text
    Next lI

    '3. Posiciona controles.
    If flxSelecao.Rows = 2 Then
        flxSelecao.Rows = 1
        cmdExcTodos.Enabled = False
        cmdExcluir.Enabled = False
        cmdOk.Enabled = False
    Else
        flxSelecao.RemoveItem flxSelecao.Row
    End If
    cmdSelTodos.Enabled = True
    cmdSelecionar.Enabled = True

    Screen.MousePointer = vbDefault
End Sub
Private Sub cmdExcTodos_Click()
    Dim lI1 As Integer      'Utilizado com For...Next.
    Dim lI2 As Byte         'Utilizado com For...Next.

    Screen.MousePointer = vbHourglass

    '1. Coloca registros da grade de resultados na grade de seleção.
    For lI1 = 1 To (flxSelecao.Rows - 1)
        flxResultado.Rows = flxResultado.Rows + 1
        flxResultado.Row = flxResultado.Rows - 1
        flxSelecao.Row = lI1
        For lI2 = 0 To 6
            flxSelecao.Col = lI2
            flxResultado.Col = lI2
            flxResultado.Text = flxSelecao.Text
        Next lI2
    Next lI1
    flxSelecao.Rows = 1

    '3. Posiciona controles.
    cmdSelTodos.Enabled = True
    cmdSelecionar.Enabled = True
    cmdExcTodos.Enabled = False
    cmdExcluir.Enabled = False
    cmdOk.Enabled = False

    Screen.MousePointer = vbDefault
End Sub
Private Sub cmdOK_Click()

    Dim clsA46V101 As clsA46V101A    'Classe: cria arquivo de transmissão.
    Dim lCodCorretor As Long      'Código do corretor.
    Dim lCodUsuTransmissao As String  'Long             'Código do usuário de transmissão.
    Dim lEndArqTransmissao As String          'Endereço do arquivo de transmissão.
    Dim lI As Integer    'Utilizado com For...Next.
    Dim lMensagem As String    'Mensagem.
    Dim lNosNumero As String  'Nosso número.
    Dim lSYASMULTINI As String    'Endereço do arquivo SYASMULT.INI.
    Dim lColecaoSelecionada As Collection       'Coleção de pedidos selecionado.

    On Error GoTo ControleErro

    Screen.MousePointer = vbHourglass

    '1. Cria matriz com relação dos documentos a transmitir.
    With flxSelecao
        'a) Verifica se grade está preenchida.
        If .Rows = 1 Then   'Apenas o cabeçalho.
            Call gM46V999.gpGraLog(1, "FGE0090 - Não há documento selecionado para transmissão.")
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        'b) Posiciona dados na matriz.
        For lI = 1 To (.Rows - 1)
            'a) Posiciona nosso número.
            .Row = lI
            .Col = 5    'Nosso número.
            lNosNumero = .Text
            'b) Inclui na coleção
            Set objstcPedido = Nothing
            Set objstcPedido = objPedido.Colecao.Item(lNosNumero)
            Set lColecaoSelecionada = New Collection
            lColecaoSelecionada.Add objstcPedido
            'c) Verifica e preenche código do corretor.
            If lCodCorretor = 0 Then
                .Col = 6        'Código do corretor.
                If gM46V999.gfPreenchido(.Text) Then
                    If IsNumeric(.Text) Then
                        lCodCorretor = CLng(.Text)
                    End If
                End If
            End If

            If lCodCorretor = 0 Then
                lCodCorretor = 9999999
            End If

            'Crio o arquivo de transmissão. (1 A 1)
            Set clsA46V101 = New clsA46V101A
            Set clsA46V101.M46V999 = gM46V999

            If clsA46V101.gfCriArqTransmissao(lCodCorretor, lCodUsuTransmissao, lColecaoSelecionada, lEndArqTransmissao, lMensagem) = False _
               Then
                If gM46V999.gfPreenchido(lMensagem) Then
                    Call gM46V999.gpGraLog(1, lMensagem)
                End If
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            Set clsA46V101 = Nothing
            Call gM46V999.gpGraLog(0, "AGE0042 - " & gM46V999.gPrefixo & "Criado arquivo de transmissão " & lEndArqTransmissao & ".")
        Next lI
    End With


    '3. Transmite arquivo de proposta.
    Call gfExecTransmissao(enum_TipTransmissao_Transmissao, False)

    '4. Altera registros de dados comuns para indicar proposta transmitida.
    '*****************MARCADA COMO TRANSMITIDA SOMENTE AO RECEBER O PROTOCOLO
    '    For Each objstcPedido In lColecaoSelecionada
    '        objstcPedido.Cod_Sit_Ped = e_SitPedido_Transmitido
    '        If Not objPedido.mfAlterar(objstcPedido) Then
    '            Call gM46V999.gpGraLog(1, objPedido.UltimoErro)
    '            Screen.MousePointer = vbDefault
    '            Exit Sub
    '        End If
    '    Next objstcPedido

    Screen.MousePointer = vbDefault
    Unload Me
    Exit Sub
ControleErro:
    Call gM46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gM46V999.gpGraLog(0, "T46V005A.cmdOk_Click ERRO - " & Err & " " & Err.Description)
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSelecionar_Click()
    Dim lI As Byte        'Utilizado com For...Next.

    Screen.MousePointer = vbHourglass

    '1. Verifica se tem documento selecionado.
    If flxResultado.Row = 0 Then
        Call gM46V999.gpGraLog(1, "FNP0014 - Documento não selecionado.")
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    '2. Coloca registro selecionado na grade de seleção.
    flxSelecao.Rows = flxSelecao.Rows + 1
    flxSelecao.Row = flxSelecao.Rows - 1
    For lI = 0 To 6
        flxResultado.Col = lI
        flxSelecao.Col = lI
        flxSelecao.Text = flxResultado.Text
    Next lI

    '3. Posiciona controles.
    If flxResultado.Rows = 2 Then
        flxResultado.Rows = 1
        cmdSelTodos.Enabled = False
        cmdSelecionar.Enabled = False
    Else
        flxResultado.RemoveItem flxResultado.Row
    End If
    flxSelecao.ColSel = flxSelecao.Cols - 1

    cmdExcTodos.Enabled = True
    cmdExcluir.Enabled = True
    cmdOk.Enabled = True

    Screen.MousePointer = vbDefault
End Sub
Private Sub cmdSelTodos_Click()
    Dim lI1 As Integer      'Utilizado com For...Next.
    Dim lI2 As Byte         'Utilizado com For...Next.

    Screen.MousePointer = vbHourglass

    '1. Coloca registros da grade de resultados na grade de seleção.
    For lI1 = 1 To (flxResultado.Rows - 1)
        flxSelecao.Rows = flxSelecao.Rows + 1
        flxSelecao.Row = flxSelecao.Rows - 1
        flxResultado.Row = lI1
        For lI2 = 0 To 6
            flxResultado.Col = lI2
            flxSelecao.Col = lI2
            flxSelecao.Text = flxResultado.Text
        Next lI2
    Next lI1
    flxResultado.Rows = 1

    '3. Posiciona controles.
    cmdSelTodos.Enabled = False
    cmdSelecionar.Enabled = False
    cmdExcTodos.Enabled = True
    cmdExcluir.Enabled = True
    cmdOk.Enabled = True

    Screen.MousePointer = vbDefault
End Sub
Private Sub flxResultado_DblClick()
    Call cmdSelecionar_Click
End Sub
Private Sub flxSelecao_DblClick()
    Call cmdExcluir_Click
End Sub
Private Sub Form_Activate()
    frmT46V001A.panForm = Mid$(Me.Name, 4)

    If fraSelecao.Visible = False Then
        Unload Me
    End If
End Sub
Private Sub Form_Load()
    Dim lMensagem As String         'Mensagem.
    Dim lObjstcCorr As stcA46V705B      'Objeto da classe de estrutura stcA46V705B (TAB_PED_CORR).


    On Error GoTo ControleErro

    '1. Centraliza form.
    Me.Height = 6000
    Call gM46V999.gpCenForm(Me)

    '2. Posiciona controles.
    lblUsuario.Caption = gM46V999.gNomUsuario
    cmdExcTodos.Enabled = False
    cmdExcluir.Enabled = False
    cmdOk.Enabled = False

    '3. Instancio o objeto.
    Set objPedido = New clsA46V702A
    Set objPedido.M46V999 = gM46V999

    Set objstcPedido = New stcA46V702B
    Set objstcPedido.M46V999 = gM46V999

    objstcPedido.Cod_Sit_Ped = e_SitPedido_ComCalculo
    objstcPedido.Cod_Tip_Doc = e_TipDoc_Proposta

    '4. Busco os pedidos.
    If Not objPedido.mfCarregarColecao(objstcPedido, "", "", True) Then
        Call gM46V999.gpGraLog(1, objPedido.UltimoErro)
        Exit Sub
    End If

    '5. Se não encontrou registro, desabilita controles.
    If objPedido.Colecao.Count = 0 Then
        fraSelecao.Visible = False
        cmdSelTodos.Visible = False
        cmdSelecionar.Visible = False
        cmdExcTodos.Visible = False
        cmdExcluir.Visible = False
        flxResultado.Visible = False
        Call gM46V999.gpGraLog(1, "AGE0037 - Não há proposta pronta para transmissão.")
        Exit Sub
    End If

    '6. Monta grades.
    Call mpMonGrade(flxResultado)
    Call mpMonGrade(flxSelecao)

    '7. Coloca registros na grade.
    With flxResultado
        For Each objstcPedido In objPedido.Colecao
            'Ficha 901620
            If objstcPedido.Dat_Pri_Calculo > 0 Then
                If DateDiff("D", gM46V999.gfForData(2, objstcPedido.Dat_Pri_Calculo), Date) <= 35 Then
                    If DateDiff("D", gM46V999.gfForData(2, objstcPedido.Dat_Liber), Date) <= 15 Then
                        If objstcPedido.Tip_Cust_Emis = 5 And objstcPedido.objstcCtrlEmis.Num_Bloq = 0 Then
                            GoTo Proximo
                        End If
                        'a) Acrescenta linha na grade.
                        .Rows = .Rows + 1
                        'Pedido
                        .TextMatrix(.Rows - 1, 0) = Format$(objstcPedido.Num_Ped, "000000")
                        'Segurado.
                        .TextMatrix(.Rows - 1, 1) = objstcPedido.Nom_Segurado
                        'CNPJ/CPF.
                        .TextMatrix(.Rows - 1, 2) = objstcPedido.Num_Cgc_Cpf
                        'Início de vigência.
                        .TextMatrix(.Rows - 1, 3) = gM46V999.gfForData(2, objstcPedido.Dat_Ini_Vig)
                        'Término de vigência.
                        .TextMatrix(.Rows - 1, 4) = gM46V999.gfForData(2, objstcPedido.Dat_Ter_Vig)
                        'Nosso número.
                        .TextMatrix(.Rows - 1, 5) = objstcPedido.NOSSO_NUMERO
                        'Corretor
                        If objstcPedido.ColecaoPedCorr.Count > 0 Then
                            Set lObjstcCorr = objstcPedido.ColecaoPedCorr.Item(1)
                            .TextMatrix(.Rows - 1, 6) = lObjstcCorr.Cod_Corr
                        Else
                            .TextMatrix(.Rows - 1, 6) = 0
                        End If
                        If .Row > 15 Then
                            .TopRow = .Row - 14
                        End If
                        .Refresh
                    End If
                End If
            End If
Proximo:
        Next objstcPedido

        If .Rows = 1 Then
            fraSelecao.Visible = False
            cmdSelTodos.Visible = False
            cmdSelecionar.Visible = False
            cmdExcTodos.Visible = False
            cmdExcluir.Visible = False
            flxResultado.Visible = False
            Call gM46V999.gpGraLog(1, "AGE0037 - Não há proposta pronta para transmissão.")
            Exit Sub
        End If
        .TopRow = 1
        .Refresh
        .ColSel = .Cols - 1
        .Row = 1
    End With
    Exit Sub
ControleErro:
    Call gM46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gM46V999.gpGraLog(0, "T46V005A.Load ERRO - " & Err & " " & Err.Description)
End Sub

