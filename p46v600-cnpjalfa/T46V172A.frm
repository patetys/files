VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT46V172A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Solicitação de cotação"
   ClientHeight    =   4275
   ClientLeft      =   2550
   ClientTop       =   1860
   ClientWidth     =   7800
   Icon            =   "T46V172A.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtObservacao 
      ForeColor       =   &H00C00000&
      Height          =   1695
      Left            =   1920
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   1920
      Width           =   5655
   End
   Begin VB.TextBox txtNome 
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1935
      MaxLength       =   100
      TabIndex        =   1
      Top             =   1200
      Width           =   5715
   End
   Begin VB.TextBox txtEmail 
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1935
      MaxLength       =   50
      TabIndex        =   0
      Top             =   840
      Width           =   5715
   End
   Begin VB.TextBox txtTelefone 
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1920
      MaxLength       =   20
      TabIndex        =   2
      Top             =   1560
      Width           =   2355
   End
   Begin Threed.SSPanel panMensagem 
      Height          =   585
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   7560
      _Version        =   65536
      _ExtentX        =   13335
      _ExtentY        =   1032
      _StockProps     =   15
      Caption         =   "Se deseja solicitar esta cotação, preencha os dados abaixo e clique em 'OK'"
      ForeColor       =   12582912
      BackColor       =   12632319
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
   End
   Begin Threed.SSCommand cmdOk 
      Height          =   300
      Left            =   2760
      TabIndex        =   3
      Top             =   3840
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
      Top             =   3840
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1764
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Cancelar"
      ForeColor       =   12582912
   End
   Begin VB.Label lblPesquisa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Observação: "
      ForeColor       =   &H00C00000&
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   1920
      Width           =   1635
   End
   Begin VB.Label lblPesquisa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nome: "
      ForeColor       =   &H00C00000&
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   8
      Top             =   1200
      Width           =   1635
   End
   Begin VB.Label lblPesquisa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Email: "
      ForeColor       =   &H00C00000&
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   840
      Width           =   1635
   End
   Begin VB.Label lblPesquisa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Telefone: "
      ForeColor       =   &H00C00000&
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   1635
   End
End
Attribute VB_Name = "frmT46V172A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mAcao As Integer
Private mColecao As Collection
Private mMensagemErros As String

Public Sub gpInstanciarForm(ByRef pAcao As Integer, pColecao As Collection, pMensagem As String)
    mAcao = 0
    Set mColecao = pColecao
    mMensagemErros = pMensagem
    Me.Show vbModal
    pAcao = mAcao
End Sub

Private Sub cmdCancelar_Click()
    mAcao = 0
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim lMensagem As String
    Dim lNovNosNumero As String        'Novo: Nosso Numero.
    Dim lNovNumPedido As String        'Novo: número do pediro.
    Dim objPedidoNovo As clsA46V702A    'Objeto da classe clsA46V702A  (TAB_PED)
    Dim objstcPedidoNovo As stcA46V702B      'Objeto da classe de estrutura stcA46V702B (TAB_PED)

    If Not gM46V999.gfPreenchido(txtEmail.Text) Then
        Call gM46V999.gpGraLog(1, "FNP0050 - Email não preenchido.")
        txtEmail.SetFocus
        Exit Sub
    End If
    If Not txtEmail.Text Like "*@*" Then
        Call gM46V999.gpGraLog(1, "AGE0160 - Email inválido.")
        txtEmail.SetFocus
        Exit Sub
    End If
    If Right(txtEmail.Text, 1) = "@" Then
        Call gM46V999.gpGraLog(1, "AGE0160 - Email inválido.")
        txtEmail.SetFocus
        Exit Sub
    End If
    If Not gM46V999.gfPreenchido(txtNome.Text) Then
        Call gM46V999.gpGraLog(1, "FNP0051 - Nome não preenchido.")
        txtNome.SetFocus
        Exit Sub
    End If
    If Not gM46V999.gfPreenchido(txtTelefone.Text) Then
        Call gM46V999.gpGraLog(1, "FNP0052 - Telefone não preenchido.")
        txtTelefone.SetFocus
        Exit Sub
    End If

    'Ficha 901718
    If frmT46V101A.objstcPedido.Dat_Pri_Calculo <> 0 Then
        'Calcula nosso número e número do pedido.
        If gM46V999.gfCalNosNumero(lNovNosNumero, lMensagem) = False Then
            Call gM46V999.gpGraLog(2, lMensagem)
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        If gM46V999.gfCalNumPedido(lNovNumPedido, lMensagem) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        'Copia documento.
        gM46V999.gOrigemChamada = e_Origem_Kit
        If gM46V999.gfCopProposta(frmT46V101A.objstcPedido.NOSSO_NUMERO, lNovNosNumero, lNovNumPedido, lMensagem, 0, 0, frmT46V101A.objstcPedido.Cod_Ramo, 0, True) = False Then
            Call gM46V999.gpGraLog(2, lMensagem)
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        gM46V111.gComNosNumero = lNovNosNumero

        '[Instanciar a classe]
        Set objPedidoNovo = New clsA46V702A
        Set objPedidoNovo.M46V999 = gM46V999

        Set objstcPedidoNovo = New stcA46V702B
        Set objstcPedidoNovo.M46V999 = gM46V999


        objstcPedidoNovo.NOSSO_NUMERO = lNovNosNumero

        If Not objPedidoNovo.mfCarregarColecao(objstcPedidoNovo, "", "", True) Then
            Call gM46V999.gpGraLog(1, objPedidoNovo.UltimoErro)
            Screen.MousePointer = vbHourglass
            Exit Sub
        End If

        Set mColecao = objPedidoNovo.Colecao

    End If

    If Not mfSolicitarCotacao(gM46V111.gComNosNumero, lMensagem) Then
        Call gM46V999.gpGraLog(1, lMensagem)
        Exit Sub
    End If

    'Gerar o arquivo
    mAcao = 1
    Call gfExecTransmissao(enum_TipTransmissao_Cotacao, True)
    frmT46V101A.panAguarde.Caption = "Esta cotação foi solicitada a Yasuda e está pendente em processo de análise, nenhuma alteração poderá ser feita."
    frmT46V101A.panAguarde.Visible = True
    Call gM46V999.gpGraLog(1, "Solicitação de cotação realizada." & vbCrLf & "Para consultar a situação da cotação solicitada, acesse a Área do Corretor - YasudaNet e, no menu à esquerda da tela, selecione a opção 'Consulta diversos' do item 'Cotações'")
    Unload Me
End Sub

Private Function mfSolicitarCotacao(pNossoNumero As String, pMensagem As String) As Boolean
    Dim lobjstcTransm As stcA28V721B
    Dim lobjTransm As clsA28V721A
    Dim lUsuario As Long
    Dim lCodCorretor As Long
    Dim lArqIni As String
    Dim clsA46V101A As clsA46V101A
    Dim lObservacao As String

    Dim lPropostas(1) As String         'Matriz com propostas selecionadas (neste caso, apenas 1).
    Dim lArqTransmissao As String

    On Error GoTo ControleErro

    Set lobjstcTransm = New stcA28V721B
    Set lobjTransm = New clsA28V721A
    Set lobjTransm.M46V999 = gM46V999

    lobjstcTransm.NOSSO_NUMERO = pNossoNumero
    lobjstcTransm.STATUS = -1

    If Not lobjTransm.gfCarregarColecao(lobjstcTransm, "") Then
        pMensagem = lobjTransm.UltimoErro
        Exit Function
    End If

    If lobjTransm.Colecao.Count > 0 Then
        Set lobjstcTransm = lobjTransm.Colecao.Item(1)
    End If

    With lobjstcTransm
        .CCotacao = ""
        .Dat_Cotacao = Left(gM46V999.gfForData(3), 8)
        .Email = txtEmail.Text
        .Nome = txtNome.Text
        .Telefone = txtTelefone.Text
        .NOSSO_NUMERO = pNossoNumero
        .STATUS = 0
        .Usuario = gM46V999.gNomUsuario
    End With

    If lobjTransm.Colecao.Count > 0 Then
        If Not lobjTransm.gfAlterar(lobjstcTransm, "") Then
            pMensagem = lobjTransm.UltimoErro
            Exit Function
        End If
    Else
        If Not lobjTransm.gfIncluir(lobjstcTransm, "") Then
            pMensagem = lobjTransm.UltimoErro
            Exit Function
        End If
    End If
    lPropostas(1) = pNossoNumero
    If gCodCorrInst <> 0 Then
        lCodCorretor = gCodCorrInst
    Else
        lCodCorretor = 9999999
    End If
    lObservacao = txtObservacao.Text
    lObservacao = Replace(lObservacao, vbCrLf, "%!*")

    lObservacao = "** Erros classe 4: " & Replace(mMensagemErros, vbCrLf, "%!*") & "%!*" & " ** Observação do corretor: " & "%!*" & lObservacao
    Set clsA46V101A = New clsA46V101A
    Set clsA46V101A.M46V999 = gM46V999

    If clsA46V101A.gfCriArqCotacao(lCodCorretor, gM46V999.gNomUsuario, _
                                   mColecao, txtEmail.Text, txtNome.Text, txtTelefone.Text, lobjstcTransm.CCotacaoOrig, lArqTransmissao, pMensagem, lObservacao) = False Then
        Call gM46V999.gpGraLog(1, pMensagem)
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set clsA46V101A = Nothing


    mfSolicitarCotacao = True

    Exit Function
ControleErro:
    If Not gM46V999.gfPreenchido(pMensagem) Then
        pMensagem = Err.Description
    End If
End Function

Private Sub Form_Load()
    Call gM46V999.gpCenForm(Me)
End Sub


Private Sub txtEmail_GotFocus()
    With txtEmail
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub
Private Sub txtEmail_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtNome_GotFocus()
    With txtNome
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub
Private Sub txtNome_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtObservacao_LostFocus()
    txtObservacao.Text = gM46V999.gfRetirarCaracterEspecial(txtObservacao.Text, False, False)
End Sub

Private Sub txtObservacao_GotFocus()
    With txtObservacao
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub
Private Sub txtObservacao_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtTelefone_GotFocus()
    With txtTelefone
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub
Private Sub txtTelefone_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
