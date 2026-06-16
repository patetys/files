VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmT46V012A 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pesquisa Proposta"
   ClientHeight    =   3045
   ClientLeft      =   9810
   ClientTop       =   3150
   ClientWidth     =   4020
   ControlBox      =   0   'False
   Icon            =   "T46V012A.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   3045
   ScaleWidth      =   4020
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraOpcaoPesquisa 
      Caption         =   "Crit�rios de pesquisa"
      ForeColor       =   &H00800000&
      Height          =   2505
      Left            =   0
      TabIndex        =   7
      Top             =   480
      Width           =   3975
      Begin VB.TextBox txtMTVVD 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1845
         MaxLength       =   6
         TabIndex        =   5
         Top             =   2120
         Width           =   1995
      End
      Begin VB.TextBox txtDV 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1845
         MaxLength       =   6
         TabIndex        =   4
         Top             =   1755
         Width           =   1995
      End
      Begin VB.ComboBox cboOpcao 
         ForeColor       =   &H00800000&
         Height          =   315
         ItemData        =   "T46V012A.frx":030A
         Left            =   150
         List            =   "T46V012A.frx":031D
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   270
         Width           =   1635
      End
      Begin VB.TextBox txtPropostaSyas 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1845
         MaxLength       =   20
         TabIndex        =   3
         Top             =   1391
         Width           =   1995
      End
      Begin VB.TextBox txtNumeroPI 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   0
         Top             =   663
         Width           =   1350
      End
      Begin VB.TextBox txtDataRecebimento 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1027
         Width           =   1995
      End
      Begin VB.TextBox txtDigito 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3390
         MaxLength       =   1
         TabIndex        =   1
         Top             =   663
         Width           =   435
      End
      Begin VB.Label lblPesquisa 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MT VVD:"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   1
         Left            =   150
         TabIndex        =   15
         Top             =   2120
         Width           =   1635
      End
      Begin VB.Label lblPesquisa 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "D�gito verificador:"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   3
         Left            =   150
         TabIndex        =   13
         Top             =   1755
         Width           =   1635
      End
      Begin VB.Label lblPesquisa 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Proposta-Syas:"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   2
         Left            =   150
         TabIndex        =   11
         Top             =   1389
         Width           =   1635
      End
      Begin VB.Label lblPesquisa 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PI: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   150
         TabIndex        =   10
         Top             =   663
         Width           =   1635
      End
      Begin VB.Label lblPesquisa 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Data de recebimento:"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   4
         Left            =   150
         TabIndex        =   9
         Top             =   1026
         Width           =   1635
      End
      Begin VB.Label Label1 
         Caption         =   "-"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   3270
         TabIndex        =   8
         Top             =   663
         Width           =   135
      End
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "cmdSair"
      Height          =   195
      Left            =   240
      TabIndex        =   14
      Top             =   840
      Width           =   375
   End
   Begin MSComctlLib.ImageList lstImagens 
      Left            =   2880
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V012A.frx":034B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V012A.frx":0665
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V012A.frx":0AB7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V012A.frx":0DD1
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V012A.frx":1223
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V012A.frx":153D
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V012A.frx":1857
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V012A.frx":1B71
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V012A.frx":1E8B
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V012A.frx":21A5
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V012A.frx":24BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V012A.frx":27D9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar barFerramentas 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   4020
      _ExtentX        =   7091
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "lstImagens"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Pesquisar"
            Object.ToolTipText     =   "Pesquisar documentos - F3."
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Relatorio"
            Object.ToolTipText     =   "Relat�rio"
            ImageIndex      =   12
            Object.Width           =   3100
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   2700
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair da pesquisa"
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Menu mnuArquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu mnuArqPesquisar 
         Caption         =   "&Pesquisar"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuTraco 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSair 
         Caption         =   "&Sair"
         Shortcut        =   ^R
      End
   End
End
Attribute VB_Name = "frmT46V012A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mAntPI As String    'Foi efetuada a busca com o PI?
Private mLinha As Long
Private mQtdErros As Long
Private mComUltData As Double
Private mComUltUsuario As String
Private mNomeComputador As String
Private mErroPremio As Boolean
Private mErroSiscad As Boolean
Private mobjstcPed As stcA46V702B    'Objeto da classe de estrutura stcA46V702B (TAB_PED).
Private lobjTriagem As clsA46v729A
Private lobjstcTriagem As stcA46v729B
Private clsA99V001 As clsA99V001
Private mUltUsuario As String
Private mUltDatAlt As Double
Private mNossoNumeroTransm As String

Private TemOutroFormAberto As Boolean


Function mfGeraCopia(pAntNosNumero As String, pNumPI As Long, pDatRecCia As Long, ByRef pNovNosNumero As String, ByRef pMensagem As String) As Boolean
    On Error GoTo ControleErro
    Dim bdBasDados As ADODB.Connection    'Base de dados.

    Dim lNovNumPedido As String     'Novo: n�mero do pedido.
    Dim lInsert As String
    Dim lIndCalcLMI As Integer

    mfGeraCopia = False
    gM46V999.gPrefixo = " Pedido " & pAntNosNumero & ". "

    pMensagem = "AGE0009 - " & gM46V999.gPrefixo & "Confirma c�pia deste documento (Rotina de c�pia de cota��o)?"
    Call gM46V999.gpGraLog(0, pMensagem & " SIM.")

    'Calcula nosso n�mero e n�mero do pedido.
    If gM46V999.gfCalNosNumero(pNovNosNumero, pMensagem) = False Then
        Exit Function
    End If

    If gM46V999.gfCalNumPedido(lNovNumPedido, pMensagem) = False Then
        Exit Function
    End If


    'Copia documento.
    gM46V999.gOrigemChamada = e_Origem_Kit
    If gM46V999.gfCopProposta(pAntNosNumero, pNovNosNumero, lNovNumPedido, pMensagem, pNumPI, pDatRecCia, , lIndCalcLMI, True) = False Then
        Exit Function
    End If

    Call gM46V999.gpGraLog(0, "AGE0002 - Processo conclu�do: c�pia realizada para pedido " & lNovNumPedido & _
                            " (nosso n�mero " & pNovNosNumero & ").")

    If gM46V999.gfAbrBasDados("P00MPED", gM46V999.gTipP00MPED, gM46V999.gEndP00MPED, bdBasDados, _
                              pMensagem) = False Then Exit Function

    lInsert = " Insert into Tab_Ped_Num_Copia values ("
    lInsert = lInsert & "'" & pNovNosNumero & "',"
    lInsert = lInsert & "'" & pAntNosNumero & "',"
    lInsert = lInsert & Left(gM46V999.gfForData(3), 8) & ")"

    If gM46V999.gfExeSQL(bdBasDados, lInsert, pMensagem) <> 0 Then
        Exit Function
    End If

    If lIndCalcLMI = 1 Then

        lInsert = " EXEC  RAMOSDIVERSOS.DBO.proc_CopiarLMI '" & pAntNosNumero & "','" & pNovNosNumero & "'," & lNovNumPedido & "," & pNumPI
        If gM46V999.gfExeSQL(bdBasDados, lInsert, pMensagem) <> 0 Then
            Exit Function
        End If
    End If

    mfGeraCopia = True
    Exit Function
ControleErro:
    pMensagem = Err & " " & Err.Description
    Call gM46V999.gpGraLog(0, "mfGeraCopia ERRO - " & pMensagem)
End Function
Private Function mfRetornarNumero(pEntrada As String) As String
    Dim lI As Long
    Dim lAuxiliar As String
    Dim lAux As String
    For lI = 1 To Len(pEntrada)
        lAux = Mid(pEntrada, lI, 1)
        Select Case Asc(lAux)
        Case 48 To 57
            lAuxiliar = lAuxiliar & lAux
        End Select
    Next lI
    If Not gM46V999.gfPreenchido(lAuxiliar) Then
        lAuxiliar = "0"
    End If

    mfRetornarNumero = lAuxiliar
End Function

Private Sub mpCapturaErros(pPref As String)
    Dim lBDP00Multi As ADODB.Connection
    Dim lRstErro As clsYasRecordSet
    Dim lMensagem As String
    Dim lSelect As String
    Dim lDelete As String

    If gM46V999.gfAbrBasDados("P00MULTI2", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lBDP00Multi, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If

    lSelect = " Select * from Tab_Erro"

    If gM46V999.gfObtRegistro(lBDP00Multi, lSelect, lRstErro, lMensagem) = False Then
        Call mpInformarErro(lMensagem, 2, True)
        Exit Sub
    End If

    Do While Not lRstErro.EOF
        If Trim(lRstErro("DSC_ERRO")) = "PARA COBERTURA 0032, FALTA A COBERTURA 0039" And gM46V999.gCodProduto >= 938 Then
            GoTo LBL_FIM
        End If

        If Trim(lRstErro("DSC_ERRO")) = "PARA COBERTURA 0024, A IS NAO PODE SER SUPERIOR A  20,00% DA COB. 0001" And gM46V999.gCodProduto >= 950 And gM46V999.gCod_Corretor = 15701 Then
            GoTo LBL_FIM
        End If
        If gM46V999.gfPreenchido(mobjstcPed.objstcCtrlEmis.Num_Cotac) Then GoTo LBL_FIM

        Call mpInformarErro(pPref & " - " & Trim(lRstErro("DSC_ERRO")), 2, True)
LBL_FIM:
        lRstErro.MoveNext
    Loop

    lDelete = " delete from Tab_Erro"

    If gM46V999.gfExeSQL(lBDP00Multi, lDelete, lMensagem) <> 0 Then
        Call mpInformarErro(lMensagem, 2, True)
        Exit Sub
    End If

    Call gM46V999.gpFecha2(lBDP00Multi)
End Sub



Private Sub mpConsisteComum(pobjstcPed As stcA46V702B, pobjstcPedLoc As stcA46V708B, pCod_Retorno As Integer)
    Dim lretorno As String

    lretorno = gM46V111.gfExecDLL(pobjstcPed, pobjstcPedLoc, e_TipProc_ConsComum, pobjstcPedLoc.Num_Item, pCod_Retorno, gM46V999.gTipoHomologacao)

    If InStr(1, lretorno, "Existem erros de consist�ncia!") <> 0 Then
        'Busca os erros de consist�ncia.
        Call mpCapturaErros("Consist�ncia do comum")
    ElseIf gM46V999.gfPreenchido(lretorno) = True Then    'retornou sem erros
        Call mpInformarErro("Consist�ncia do comum - " & lretorno, 2, True)
    End If

    DoEvents

End Sub

Private Sub mpEmiMensagens(ByVal pPrefixo As String)
'Procedure: emite mensagens (transfere da matriz para a grade de mensagens).

'Par�metro de entrada...pPrefixo...Prefixo da mensagem.

    Dim lI As Integer    'Utilizado com For...Next.

'1. Posiciona mensagens.
    For lI = 1 To gM46V999.gQuaErros
        Call mpInformarErro(CStr(gM46V999.gerros(lI)), 2, True)
    Next lI

    gM46V999.gQuaErros = 0

End Sub

Private Sub mpGravarTriagem(pStatus As e_status_triagem)

    On Error GoTo ControleErro

    Set lobjstcTriagem = New stcA46v729B

    With lobjstcTriagem
        .NUM_PI = txtNumeroPI.Text
        .Dat_Triagem = gM46V999.gfForData(3)
        .NOSSO_NUMERO = mNossoNumeroTransm
        .STATUS = pStatus
        .User_Triagem = gM46V999.gNomUsuario6
    End With

    Set lobjTriagem = New clsA46v729A
    Set lobjTriagem.M46V999 = gM46V999

    If Not lobjTriagem.gfIncluir(lobjstcTriagem) Then
        Call gM46V999.gpGraLog(1, lobjTriagem.UltimoErro)
        Exit Sub
    End If

    Exit Sub

ControleErro:
    Call gM46V999.gpGraLog(1, Err.Description)

End Sub

Private Sub mpInformarErro(ByVal pTexto As String, pStatus As Byte, pDeslocado As Boolean)
    Dim lPos As Long  'Posi��o de caracter.
    Dim lTamanho As Long      'Tamanho.
    Dim lAuxiliar As String     'Auxiliar.

    If pDeslocado = True Then
        lPos = 1
        lTamanho = 105
        Do
            Do
                lAuxiliar = Mid(" - " & pTexto, lPos, lTamanho)

                If Right(lAuxiliar, 1) <> " " And Len(lAuxiliar) = lTamanho Then
                    lTamanho = lTamanho - 1
                Else
                    Exit Do
                End If
            Loop

            If gM46V999.gfPreenchido(lAuxiliar) Then
                lPos = lPos + lTamanho
                lTamanho = 105
                frmT46V012B.flxResultado.Row = mLinha
                frmT46V012B.flxResultado.Col = 0
                mLinha = mLinha + 1
                frmT46V012B.flxResultado.AddItem "", mLinha
                frmT46V012B.flxResultado.TextMatrix(mLinha, 1) = Space(10) & lAuxiliar
                If frmT46V012B.flxResultado.CellPicture <> 0 Then
                    Set frmT46V012B.flxResultado.CellPicture = frmT46V012B.picChecked.Picture
                End If
            Else
                Exit Do
            End If
        Loop

        mQtdErros = mQtdErros + 1
    Else
        mLinha = mLinha + 1

        frmT46V012B.flxResultado.Col = 0
        frmT46V012B.flxResultado.Row = mLinha
        frmT46V012B.flxResultado.CellPictureAlignment = 4

        If pStatus = 1 Then
            Set frmT46V012B.flxResultado.CellPicture = frmT46V012B.picOk.Picture
        Else
            Set frmT46V012B.flxResultado.CellPicture = frmT46V012B.picChecked.Picture
        End If
    End If

End Sub

Private Sub barFerramentas_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case UCase(Button.Key)
    Case "PESQUISAR"
        Call mnuArqPesquisar_Click
    Case "RELATORIO"
        Call mnuExiRelatorio_Click
    Case "SAIR"
        Call mnuSair_Click
    End Select
End Sub



Private Sub cboOpcao_Click()
'Ficha 1000449 - Incluir campo Referencia Marsh para digita��o
    If cboOpcao.ListIndex = 2 Or cboOpcao.ListIndex = 3 Then
        fraOpcaoPesquisa.Height = 2505
        Me.Height = 3735
        'Ficha 1000449 - Incluir campo Referencia Marsh para digita��o
        If cboOpcao.ListIndex = 2 Then
            lblPesquisa(1).Caption = "MT VVD:"
        Else
            lblPesquisa(1).Caption = "Referencia Marsh:"
        End If
    Else
        fraOpcaoPesquisa.Height = 2100
        Me.Height = 3250
    End If
End Sub

Private Sub cmdSair_Click()
    Call mnuSair_Click
End Sub

Private Sub Form_Load()
    Dim lDigito As String
    Dim lMensagem As String
    Dim bdBasDados As ADODB.Connection    'Base de dados.
    Dim lretorno As Integer
    Dim lCodigoUnic As Long
    Dim lSql As String
    Dim lrst As clsYasRecordSet

    TemOutroFormAberto = False

    mAntPI = ""
    Me.Height = 3285
    Call gM46V999.gpCenForm(Me)  'Centraliza form.
    ' If gM46V999.PerfilAcesso <> e_Perfil_DIG_IST Then
    '     txtDataRecebimento.Text = GetSetting(App.EXEName, "DataRecebimento", Me.Caption, "")
    ' End If
    cboOpcao.ListIndex = 0
    txtDataRecebimento.Text = ""
    txtDigito.Text = ""
    txtDV.Text = ""
    txtMTVVD.Text = ""
    txtNumeroPI.Text = ""
    txtPropostaSyas.Text = ""

    If gM46V999.PerfilAcesso <> e_Perfil_Triagem Then
        barFerramentas.Buttons.Item(3).Enabled = False
    Else
        gM46V111.gComNosNumero = ""
        Call gM46V111.gpObterRegComunicacao(False, "", "")
        If gM46V999.gfPreenchido(gM46V111.gComNosNumero) Then
            If gM46V111.gNum_Pi <> 0 Then
                txtNumeroPI.Text = gM46V111.gNum_Pi
                Call txtNumeroPI_LostFocus
                mAntPI = txtNumeroPI.Text
                Set clsA99V001 = New clsA99V001
                Set clsA99V001.M46V999 = gM46V999

                If clsA99V001.DIG_YASUDA(txtNumeroPI.Text, lDigito, lMensagem) = False Then
                    Exit Sub
                End If
                Set clsA99V001 = Nothing
                txtDigito.Text = Format(lDigito, "0")
                If gM46V999.gfPreenchido(txtPropostaSyas.Text) = False Then
                    txtPropostaSyas.Text = gM46V111.gComNosNumero

                    If gM46V999.gfAbrBasDados("P00MPED", gM46V999.gTipP00MPED, gM46V999.gEndP00MPED, bdBasDados, _
                                              lMensagem) = False Then
                        Call gM46V999.gpGraLog(1, "FNP0044 - " & lMensagem)
                        Exit Sub
                    End If

                    lSql = " Select num_dv from ged.dbo.tab_triagem where num_pi = " & gM46V111.gNum_Pi
                    If Not gM46V999.gfObtRegistro(bdBasDados, lSql, lrst, lMensagem) Then
                        Call gM46V999.gpGraLog(1, "FNP0044 - " & lMensagem)
                        Exit Sub
                    End If
                    lCodigoUnic = 0
                    If Not lrst.EOF Then
                        lCodigoUnic = lrst("Num_DV")
                    End If




                    Call mpVerificaUnicidade(bdBasDados, gM46V111.gComNosNumero, lCodigoUnic, lretorno, lMensagem, lCodigoUnic)
                    txtDV.Text = lCodigoUnic
                End If
            End If
        End If
    End If
End Sub


Private Sub txtDataRecebimento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtPropostaSyas.Enabled = False Then
            Call mnuArqPesquisar_Click
        Else
            SendKeys vbTab
        End If
    End If
End Sub

Private Sub txtDataRecebimento_LostFocus()
    If TemOutroFormAberto = True Then Exit Sub

    If txtDataRecebimento.Text <> "" Then
        txtDataRecebimento.Text = gM46V999.gfForData(9, txtDataRecebimento.Text)
    End If

End Sub

Private Sub mnuArqPesquisar_Click()
    Screen.MousePointer = vbHourglass

    If gM46V999.PerfilAcesso = e_Perfil_Triagem Then
        Call mpPesquisaTriagem
    Else
        Call mpPesquisaDigitacao
    End If

    Screen.MousePointer = vbDefault
End Sub

Private Sub mpPesquisaTriagem()

    Dim lMensagem As String     'Mensagem
    Dim lNossoNumero As String        'Nosso Numero.
    Dim lNovNosNumero As String         'Novo: Nosso Numero.
    Dim lNovNumPedido As String         'Novo: n�mero do pediro.
    Dim lControl As Integer
    Dim lobjCtrlEmis As clsA46V712A   'Objeto da classe clsA46V712A (TAB_CTRL_EMIS)
    Dim lobjstcCtrlEmis As stcA46V712B      'Objeto da classe de estrutura stcA46V712B (TAB_CTRL_EMIS)
    Dim bdBasDados As ADODB.Connection    'Base de dados.
    Dim lObjPed As clsA46V702A    'Objeto da classe clsA46V702A (TAB_PED).
    Dim lobjPedLoc As clsA46V708A    'Objeto da classe clsA46V708A (TAB_PED_LOC).
    Dim lobjstcPedLoc As stcA46V708B    'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC).
    Dim lretorno As Integer
    Dim lIncluirSiscad As Boolean
    Dim lAlterarSiscad As Boolean
    Dim lVerSiscad As Boolean
    Dim lDemaisErros As Boolean
    Dim UltimoErro As String

    If gM46V999.gfAbrBasDados("P00MPED", gM46V999.gTipP00MPED, gM46V999.gEndP00MPED, bdBasDados, _
                              lMensagem) = False Then
        Call gM46V999.gpGraLog(1, "FNP0044 - " & lMensagem)
        Exit Sub
    End If

    If Not mfConsistirDados Then Exit Sub

    lControl = 0

Retorno:

    'Tenta localizar os dados no tab_ctrl_emis pelo PI.
    'Instancio os objetos

    Set lobjCtrlEmis = New clsA46V712A
    Set lobjCtrlEmis.M46V999 = gM46V999

    Set lobjstcCtrlEmis = New stcA46V712B

    Set lObjPed = New clsA46V702A
    Set lObjPed.M46V999 = gM46V999

    Set mobjstcPed = New stcA46V702B
    Set mobjstcPed.M46V999 = gM46V999

    Set lobjCtrlEmis.ConexaoBD = bdBasDados

    If gM46V999.gfPreenchido(txtNumeroPI) Then
        lobjstcCtrlEmis.NUM_PI = txtNumeroPI

        If Not lobjCtrlEmis.mfCarregarColecao(lobjstcCtrlEmis, "", "") Then
            UltimoErro = lobjCtrlEmis.UltimoErro
            Call gM46V999.gpGraLog(1, "FNP0024 - " & UltimoErro)
            Exit Sub
        End If

        If lobjCtrlEmis.Colecao.Count <> 0 Then
            '*****Localizou o PI
            Set lobjstcCtrlEmis = lobjCtrlEmis.Colecao.Item(1)
            txtPropostaSyas = lobjstcCtrlEmis.NOSSO_NUMERO

            If lobjstcCtrlEmis.Dat_Rec_Cia < DateAdd("d", -15, Date) And gM46V999.PerfilAcesso <> e_Perfil_DIG_Gestor And lobjstcCtrlEmis.Num_Apol = 0 Then
                Call gM46V999.gpGraLog(1, "FGE0079 - Data de recebimento inv�lida (inferior a 15 dias da data de " & Date & ")")
                txtDataRecebimento.Text = lobjstcCtrlEmis.Dat_Rec_Cia

                If txtDataRecebimento.Visible = True And txtDataRecebimento.Enabled = True Then
                    txtDataRecebimento.SetFocus
                End If

                Exit Sub
            End If

            lNovNosNumero = lobjstcCtrlEmis.NOSSO_NUMERO
            GoTo RegistroLocalizado
        End If
    End If

    If CDate(frmT46V012A.txtDataRecebimento) < DateAdd("d", -15, Date) And gM46V999.PerfilAcesso <> e_Perfil_DIG_Gestor Then
        Call gM46V999.gpGraLog(1, "FGE0080 - Data de recebimento inv�lida (inferior a 15 dias da data de " & Date & ")")

        If txtDataRecebimento.Visible = True And txtDataRecebimento.Enabled = True Then
            txtDataRecebimento.SetFocus
        End If

        Exit Sub
    End If

    If gM46V999.gfPreenchido(txtPropostaSyas.Text) Then
        If Not gM46V999.gBaseLocal Then
            lNossoNumero = txtPropostaSyas.Text

            If Not mpVerificaUnicidade(bdBasDados, lNossoNumero, gM46V999.CDblx(txtDV.Text), lretorno, lMensagem) Then
                Call gM46V999.gpGraLog(2, lMensagem)

                If txtDV.Visible = True And txtDV.Enabled = True Then
                    txtDV.SetFocus
                End If

                Exit Sub
            End If

            If lretorno = 1 Then
                '****Transmiss�o n�o localizada******
                lobjstcCtrlEmis.NUM_PI = 0
                lobjstcCtrlEmis.NOSSO_NUMERO = txtPropostaSyas.Text

                If Not lobjCtrlEmis.mfCarregarColecao(lobjstcCtrlEmis, "", "") Then
                    UltimoErro = lobjCtrlEmis.UltimoErro
                    Call gM46V999.gpGraLog(1, "FNP0024 - " & UltimoErro)
                    Exit Sub
                End If

                If lobjCtrlEmis.Colecao.Count <> 0 Then
                    Set lobjstcCtrlEmis = lobjCtrlEmis.Colecao.Item(1)
                    If Not mfGeraCopia(lobjstcCtrlEmis.NOSSO_NUMERO, txtNumeroPI, gM46V999.gfForData(1, txtDataRecebimento), lNovNosNumero, lMensagem) Then
                        Call gM46V999.gpGraLog(2, lMensagem)
                        Exit Sub
                    Else
                        If gM46V111.gfGraRegComunicacao(lNovNosNumero, lMensagem) = False Then
                            Call gM46V999.gpGraLog(2, lMensagem)
                            Exit Sub
                        End If
                        Me.Hide
                        Exit Sub
                    End If
                End If

                Call gM46V999.gpGraLog(1, "FNL0041 - N�mero da proposta n�o localizado (" & txtPropostaSyas.Text & ").")
            ElseIf lretorno = 2 Then
                '****Transmiss�o localizada com DV diferente******
                Call gM46V999.gpGraLog(1, "FNL0041 - Proposta localizada com DV diferente do informado.")
            Else
                '****Transmiss�o localizada******
                'Calcula nosso n�mero e n�mero do pedido.

                If gM46V999.gfCalNosNumero(lNovNosNumero, lMensagem) = False Then
                    Call gM46V999.gpGraLog(1, "FNP0021 - Pesquisa Digita��o: " & lMensagem)
                    Exit Sub
                End If

                If gM46V999.gfCalNumPedido(lNovNumPedido, lMensagem) = False Then
                    Call gM46V999.gpGraLog(1, "FNP0022 - Pesquisa Digita��o: " & lMensagem)
                    Exit Sub
                End If

                If Not gM46V001.gfImportarTransmissao(lNossoNumero, gM46V999.CDblx(txtDV.Text), lNovNosNumero, lNovNumPedido, txtNumeroPI, gM46V999.gfForData(1, frmT46V012A.txtDataRecebimento), lMensagem) Then
                    UltimoErro = lMensagem
                    Call gM46V999.gpGraLog(1, "FNP0024 - " & UltimoErro)
                    Exit Sub
                End If

RegistroLocalizado:
                gM46V111.gNum_Pi = txtNumeroPI.Text
                Set lObjPed = New clsA46V702A
                Set lObjPed.M46V999 = gM46V999

                Set mobjstcPed = New stcA46V702B
                Set mobjstcPed.M46V999 = gM46V999

                mobjstcPed.NOSSO_NUMERO = lNovNosNumero

                If Not lObjPed.mfCarregarColecao(mobjstcPed, "", "", True) Then
                    UltimoErro = lobjCtrlEmis.UltimoErro
                    Call gM46V999.gpGraLog(1, "FNP0024 - " & UltimoErro)
                    Exit Sub
                End If

                If lObjPed.Colecao.Count <> 0 Then

                    Set mobjstcPed = lObjPed.Colecao.Item(1)
                    Set lobjstcPedLoc = mobjstcPed.ColecaoPedLoc.Item(1)
                    gM46V999.gCodProduto = mobjstcPed.Cod_Produto
                    gM46V999.gCod_Ramo = mobjstcPed.Cod_Ramo
                    Call mpObtProduto(gM46V999.gCodProduto, gM46V999.gCod_Ramo)

                    'Ficha 900740
                    gM46V111.gCod_plano = mobjstcPed.Cod_Estip
                    If gM46V111.gCod_plano >= 2001 And gM46V111.gCod_plano <= 2031 Then
                        gM46V999.gProQuaMaxParcelas = 11
                    Else
                        gM46V999.gProQuaMaxParcelas = 10
                    End If

                    mobjstcPed.Cod_Tip_Doc = e_TipDoc_Proposta

                    If gM46V999.gCod_Ramo = e_CodRamo_Empresarial Or _
                       gM46V999.gCod_Ramo = e_CodRamo_Condominio Or _
                       gM46V999.gCod_Ramo = e_CodRamo_Residencial Or _
                       gM46V999.gCod_Ramo = e_CodRamo_Riscos_diversos Or _
                       gM46V999.gCod_Ramo = e_CodRamo_Roubos Then
                        gM46V111.gTipoGrupMulti = e_GrupoRamo_Multi
                    ElseIf gM46V999.gCod_Ramo = e_CodRamo_D_P_E_M Or _
                           gM46V999.gCod_Ramo = e_CodRamo_Ap_bilhete Then
                        gM46V111.gTipoGrupMulti = e_GrupoRamo_570_820
                    ElseIf gM46V999.gCod_Ramo = e_CodRamo_Automoveis Then
                        gM46V111.gTipoGrupMulti = e_GrupoRamo_Auto
                    ElseIf gM46V999.gCod_Ramo = e_CodRamo_Acidentes_pessoais Then
                        gM46V111.gTipoGrupMulti = e_GrupoRamo_Acidentes_pessoais
                    ElseIf gM46V999.gCod_Ramo = e_CodRamo_Transp_internacional Or _
                           gM46V999.gCod_Ramo = e_CodRamo_Transp_nacional Or _
                           gM46V999.gCod_Ramo = e_CodRamo_Rc_Transportador Or _
                           gM46V999.gCod_Ramo = e_CodRamo_Rcf_dc Then
                        gM46V111.gTipoGrupMulti = e_GrupoRamo_Transp_Diversos
                    ElseIf gM46V999.gCod_Ramo <> e_CodRamo_NaoDefinido Then
                        gM46V111.gTipoGrupMulti = e_GrupoRamo_Generico
                    End If


                    'Ficha 901824
                    'Vai verificar se existem apontamentos na triagem para RELACAO DE BENS, SOMENTE PARA PREDIO e SOMENTE PARA CONTEUDO, incluindo os textos para a emiss�o do Booklet
                    'Vai verificar se o campo PRIMEIRA PARCELA ou SEGUNDA PARCELA est�o preenchidos, se sim, dever� abrir a tela de vencimentos do Syas
                    If Not mfVerDizer(mobjstcPed, lMensagem) Then
                        Call gM46V999.gpGraLog(2, lMensagem)
                        Exit Sub
                    End If


                    'Busco siscad
                    'Comparo siscad
                    If cboOpcao.ItemData(cboOpcao.ListIndex) = 3 Then
                        mobjstcPed.objstcCtrlEmis.Num_Ref_Corr = "MT" & txtMTVVD.Text
                    End If
                    'Ficha 1000449 - Incluir campo Referencia Marsh para digita��o
                    If cboOpcao.ItemData(cboOpcao.ListIndex) = 4 Then
                        mobjstcPed.objstcCtrlEmis.Num_Ref_Corr = txtMTVVD.Text
                    End If
                    If Not mfBuscarDadosSiscad(mobjstcPed.Num_Apol, _
                                               lIncluirSiscad, lAlterarSiscad, lVerSiscad, lobjstcPedLoc.Tip_Emissao, lMensagem) Then
                        Call gM46V999.gpGraLog(1, lMensagem)
                        Exit Sub
                    End If
                    mNossoNumeroTransm = " "
                    'Calcula
                    'Compara premio
                    Call mpCalcular(lNossoNumero, lobjstcPedLoc, lobjstcPedLoc.Tip_Emissao, lDemaisErros)

                    'exibe mensagem
                    'Emite ou suspende

                    If mQtdErros > 0 Then
                        gM46V999.gEfetuouCalculo = False
                    Else
                        gM46V999.gEfetuouCalculo = True
                    End If

                    If lIncluirSiscad Then
                        'Inclui o siscad
                        If Not mfIncluirSiscad(lMensagem) Then
                            Call gM46V999.gpGraLog(1, lMensagem)
                            Exit Sub
                        End If
                    End If

                    If lAlterarSiscad Then
                        'Altera siscad
                        If Not mfAlterarSiscad(lMensagem) Then
                            Call gM46V999.gpGraLog(1, lMensagem)
                            Exit Sub
                        End If
                    End If

                    If Not mfGravarEndCobranca(lMensagem) Then
                        Call gM46V999.gpGraLog(1, lMensagem)
                        Exit Sub
                    End If

                    Screen.MousePointer = vbDefault

                    Call frmT46V012B.mpInstanciarForm(mobjstcPed, lVerSiscad, _
                                                      lDemaisErros, lretorno)

                    If lretorno = 0 Then
                        'Nenhuma a��o
                        Call mpGravarTriagem(e_status_triagem_NDefinido)
                    Else
                        If lretorno = 1 Then
                            'Emitido OK
                            Call mpGravarTriagem(e_status_triagem_OK)
                        ElseIf lretorno = 2 Then
                            'Suspenso
                            If mErroPremio = True And mErroSiscad = True Then
                                Call mpGravarTriagem(e_status_triagem_Erro)
                            ElseIf mErroPremio = True Then
                                Call mpGravarTriagem(e_status_triagem_ErroPremio)
                            ElseIf mErroSiscad = True Then
                                Call mpGravarTriagem(e_status_triagem_ErroSiscad)
                            Else
                                Call mpGravarTriagem(e_status_triagem_ErroTriagem)
                            End If
                        End If
                    End If
                    Call mnuEdiLimpar_Click
                End If
            End If
        End If
    End If
End Sub

Private Sub mnuEdiLimpar_Click()

    txtNumeroPI.Text = ""
    txtDigito.Text = ""
    txtPropostaSyas.Text = ""
    txtDV.Text = ""
    txtMTVVD.Text = ""

    If txtNumeroPI.Visible = True And txtNumeroPI.Enabled = True Then
        txtNumeroPI.SetFocus
    End If
    mAntPI = ""
End Sub

Private Sub mpCalcular(pNossoNumero As String, mobjstcPedLoc As stcA46V708B, pTipEmissao As e_TipEmissao, pDemaisErros As Boolean)

'Fun��o: calcula pr�mios.
'********************Vari�veis padr�o*************************************
    Dim lCkPoint As String    'Variavel utilizada para determinar pontos onde o c�digo j� passou.
    Dim lCkStr As String    'Variavel utilizada para guardar a posi��o das variaveis at� o ultimo
'lCkPoint. (posi��o das variaveis em ordem de declara��o)
'********************Vari�veis Declaradas para o processamento************
    Dim lMensagem As String    'Mensagem.
    Dim lPrefixo As String    'Prefixo de mensagem.
    Dim lUltData As Double    'Data da �ltima atualiza��o.
    Dim lobjstcItem As stcA46V708B    'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
    Dim lColecaoTemp As Collection    'Cole��o tempor�ria de itens.
    Dim lTotPr As Single
    Dim lTotPr_Resg As Single
    Dim lTotPr_Tarif As Single
    Dim lobjstcCobert As stcA46V704B    'Objeto da classe (stc) de coberturas.
    Dim lBoleto As Long
    Dim lObjPed As clsA46V702A
    Dim lobjPedLoc As clsA46V708A
    Dim lbdRamosDiversos As ADODB.Connection    'Conex�o do Banco local das FTP_TAB_PED
    Dim lrst As clsYasRecordSet    'Registro.
    Dim lSelect As String    'Select
    Dim lValorParc As Variant
    Dim lDiferenca As Single
    Dim lobjstcPedJur As stcA46V711B
    Dim lQtdErro As Long
    Dim lobjCobert As clsA46V704A

    lCkPoint = ""
    lCkStr = ""
    Err = 0
    gM46V999.gQuaErros = 0
    lQtdErro = mQtdErros
    '1. Consit�ncia dos dados de comum
    'a) Itens.

    Call mpInformarErro("", 1, False)

    mUltUsuario = mobjstcPed.Cod_User_Liber
    mUltDatAlt = mobjstcPed.Dat_Liber

    If mobjstcPed.ColecaoPedLoc.Count = 0 Then
        Call mpInformarErro("N�o localizado registro de item", 2, True)
    End If

    ''Valida��o dos dados.
    '1000817
    If mobjstcPed.ColecaoPedLoc.Count > 1 Then
        If mobjstcPed.Num_Apol_An = 0 Then
            'Verifico se existe algum item que � renova��o
            For Each lobjstcItem In mobjstcPed.ColecaoPedLoc
                If lobjstcItem.Tip_Emissao = e_TipEmissao_RenYasuda Or _
                   lobjstcItem.Tip_Emissao = e_TipEmissao_Endosso Then
                    ''Corrigir base
                    mobjstcPed.Num_Apol_An = lobjstcItem.Num_Apol_Anterior
                    Set lObjPed = New clsA46V702A
                    Set lObjPed.M46V999 = gM46V999

                    If Not lObjPed.mfAlterar(mobjstcPed, "", "", mUltDatAlt, mUltUsuario) Then
                        Call mpInformarErro(lObjPed.UltimoErro, 2, True)
                        Exit Sub
                    End If
                    Exit For
                End If
            Next lobjstcItem
        End If
    End If
    
    lCkPoint = "1"
    'b) Exportando os dados para as bases temp
    lBoleto = mobjstcPed.objstcCtrlEmis.Num_Bloq

    If gM46V999.gCodProduto < 1080 Then 'Necessario apenas para normas COBOL
        Call gM46V999.gpGraLog(0, "exportar")
        
        If Not gfExportarParaBaseTemp(mobjstcPed.NOSSO_NUMERO, mobjstcPed.Num_Ped, lBoleto, lMensagem) Then
            Call mpInformarErro(lMensagem, 2, True)
        End If
        
        For Each lobjstcItem In mobjstcPed.ColecaoPedLoc
            If lobjstcItem.Tip_Emissao = e_TipEmissao_RenYasuda Or _
               lobjstcItem.Tip_Emissao = e_TipEmissao_Endosso Then
                If Not gfExportarDadosApolice(lobjstcItem.Num_Apol_Anterior, lMensagem) Then
                    Call mpInformarErro(lMensagem, 2, True)
                End If
                mobjstcPed.Num_Apol_An = lobjstcItem.Num_Apol_Anterior
                Exit For
            End If
        Next lobjstcItem
    End If

    lCkPoint = "2"

    '2. Consiste dados comuns - proposta.
    Call mpVerComum(mobjstcPed, mobjstcPedLoc)

    lCkPoint = "3"

    If gM46V999.gQuaErros <> 0 Then
        'Exit Sub '
    End If

    '3. Processa item.
    'Vou copiar a cole��o porque durante o c�lculo os itens sofrem altera��o
    'e s�o removidos e adicionados novamente na cole��o, ent�o para que n�o aconte�a
    'de reprocessar o mesmo item, ent�o � feita a c�pia da cole��o.

    Call mpInformarErro("", 1, False)

    lCkPoint = "4"
    Set lColecaoTemp = New Collection
    For Each lobjstcItem In mobjstcPed.ColecaoPedLoc
        lColecaoTemp.Add lobjstcItem
    Next lobjstcItem

    'VERIFICAR AS REGRAS DE CONSISTENCIAS PARA CALCULAR O BATENTE
    If mobjstcPed.Num_Apol_An > 0 And _
      (mobjstcPed.Cod_Ramo = e_CodRamo_Residencial Or (mobjstcPed.Cod_Ramo = e_CodRamo_Empresarial And mobjstcPed.Cod_Produto >= 1080)) _
      And mobjstcPed.Cod_Dizer = 0 Then
        If gM46V103.ConsisteRegraBatente(mobjstcPed, 0, True) Then
            mobjstcPed.CalBatente = True
        End If
    End If

    For Each lobjstcItem In lColecaoTemp
        lCkPoint = "5"
        If mfCalPreItem(mobjstcPed, mobjstcPedLoc, lobjstcItem) = False Then
            lPrefixo = " Item " & lobjstcItem.Num_Item & " "
            Call mpEmiMensagens(lPrefixo)
        End If
        'i) Posiciona no pr�ximo item.
    Next lobjstcItem

    lCkPoint = "7"

    '9. Se houve erro, encerra processo.
    If gM46V999.gQuaErros <> 0 Then
        '        Exit Sub
    End If

    lTotPr = 0
    lTotPr_Resg = 0
    lTotPr_Tarif = 0

    For Each lobjstcItem In mobjstcPed.ColecaoPedLoc
        If mobjstcPed.CalBatente = False And lobjstcItem.ColecaoRegraBatente.Count = 0 Then
            For Each lobjstcCobert In lobjstcItem.ColecaoPedCobert
                'RESSEGURO
                If mobjstcPed.Dat_Ini_Vig >= 20090501 Then
                    lTotPr = lTotPr + lobjstcCobert.VAL_PR
                    lTotPr_Resg = lTotPr
                    lTotPr_Tarif = lTotPr
                Else
                    lTotPr = lTotPr + lobjstcCobert.VAL_PR
                    lTotPr_Resg = lTotPr_Resg + lobjstcCobert.Val_Pr_Resg
                    lTotPr_Tarif = lTotPr_Tarif + lobjstcCobert.Val_Pr_Tarif
                End If
            Next lobjstcCobert
        Else
            'OBTEM E GRAVA OS NOVOS VALORES DA CALCULADOS PELA REGRA DE BATENTE
            For Each lobjstcCobert In lobjstcItem.ColecaoRegraBatente
                If Not lobjCobert.mfAlterarValoresPrecificacao(lobjstcCobert, "", "") Then
                    Call mpInformarErro("Falha ao gravar valores da regra de batente", 2, True)
                End If

                'RESSEGURO
                If mobjstcPed.Dat_Ini_Vig >= 20090501 Then
                    lTotPr = lTotPr + lobjstcCobert.VAL_PR
                    lTotPr_Resg = lTotPr
                    lTotPr_Tarif = lTotPr
                Else
                    lTotPr = lTotPr + lobjstcCobert.VAL_PR
                    lTotPr_Resg = lTotPr_Resg + lobjstcCobert.Val_Pr_Resg
                    lTotPr_Tarif = lTotPr_Tarif + lobjstcCobert.Val_Pr_Tarif
                End If
            Next lobjstcCobert
        End If
    Next lobjstcItem

    If lTotPr < gM46V999.gProPremioMinimo Then
        lTotPr = gM46V999.gProPremioMinimo
    End If

    If lTotPr_Resg < gM46V999.gProPremioMinimo Then
        lTotPr_Resg = gM46V999.gProPremioMinimo
    End If

    If lTotPr_Tarif < gM46V999.gProPremioMinimo Then
        lTotPr_Tarif = gM46V999.gProPremioMinimo
    End If

    'RESSEGURO
    If mobjstcPed.Dat_Ini_Vig >= 20090501 Then
        mobjstcPed.VAL_PR = lTotPr
        mobjstcPed.Val_Pr_Ref_Resg = lTotPr
        mobjstcPed.Val_Pr_Ref_Tarif = lTotPr
    Else
        mobjstcPed.VAL_PR = lTotPr
        mobjstcPed.Val_Pr_Ref_Resg = lTotPr_Resg
        mobjstcPed.Val_Pr_Ref_Tarif = lTotPr_Tarif
    End If

    Set lObjPed = New clsA46V702A
    Set lObjPed.M46V999 = gM46V999

    mobjstcPed.Cod_User_Liber = gM46V999.gNomUsuario6
    mobjstcPed.Dat_Liber = Format(Date, "yyyymmdd")

    If Not lObjPed.mfAlterar(mobjstcPed, "", "", mUltDatAlt, mUltUsuario) Then
        Call mpInformarErro(lObjPed.UltimoErro, 2, True)
        Exit Sub
    End If
    mUltUsuario = mobjstcPed.Cod_User_Liber
    mUltDatAlt = mobjstcPed.Dat_Liber

    If Not gM46V999.gfAbrBasDados("RAMOSDIVERSOS", gM46V999.gTipDIVERSOS, gM46V999.gEndDIVERSOS, lbdRamosDiversos, lMensagem) Then
        Call gM46V999.gpMenErro(lMensagem)
        Exit Sub
    End If
    mNossoNumeroTransm = Trim(txtPropostaSyas.Text)
    If txtPropostaSyas.Enabled = False Then
        lSelect = ""
        lSelect = "select nosso_numero_old from dbo.Tab_Ped_Num_Copia where nosso_numero_new = '" & txtPropostaSyas.Text & "'"

        If Not gM46V999.gfObtRegistro(lbdRamosDiversos, lSelect, lrst, lMensagem) Then
            Call gM46V999.gpMenErro(lMensagem)
            Exit Sub
        End If

        If Not lrst.EOF Then
            pNossoNumero = lrst("nosso_numero_old")
        Else
            pNossoNumero = Trim(txtPropostaSyas.Text)
        End If
        mNossoNumeroTransm = pNossoNumero
    End If

    lSelect = ""
    lSelect = "Select VAL_PR, VAL_PR_REF_RESG, VAL_PR_REF_TARIF, VAL_PR_ANTEC "
    lSelect = lSelect & "from dbo.FTP_TAB_PED where NOSSO_NUMERO = '" & pNossoNumero & "'"
    lSelect = lSelect & " and cod_unicidade =  " & txtDV.Text


    'Busca os valores dos pr�mios de transmiss�o
    If Not gM46V999.gfObtRegistro(lbdRamosDiversos, lSelect, lrst, lMensagem) Then
        Call gM46V999.gpMenErro(lMensagem)
        Exit Sub
    End If

    'Compara��o dos pr�mios
    If Not lrst.EOF Then
        If Format(mobjstcPed.VAL_PR, "standard") <> Format(lrst("VAL_PR"), "standard") Then
            Call mpInformarErro("Valor do pr�mio divergente - Transmitido = R$" & Format(lrst("VAL_PR"), "standard") & ", calculado = R$" & Format(mobjstcPed.VAL_PR, "standard"), 2, True)
        End If

        If Format(mobjstcPed.Val_Pr_Ref_Tarif, "standard") <> Format(lrst("VAL_PR_REF_TARIF"), "standard") Then
            Call mpInformarErro("Valor do pr�mio tarif�rio divergente - Transmitido = R$" & Format(lrst("VAL_PR_REF_TARIF")) & ", calculado = R$" & Format(mobjstcPed.Val_Pr_Ref_Tarif, "standard"), 2, True)
        End If
    End If

    'Calcula o fracionamento
    If Not gfCarregaPerIOF(mobjstcPed.Dat_Ini_Vig, lMensagem) Then
        Call mpInformarErro("Problemas ao tentar localizar o IOF." & lMensagem, 2, True)
    End If

    With mobjstcPed
        Call gM46V998.gfCalFracionamento(mobjstcPed, mobjstcPedLoc.Tip_Emissao, _
                                         lValorParc, lMensagem, gM46V111.gProJurIOF, .Cod_Reg, .Val_IOF, .Val_Parc_Min)
    End With

    If mobjstcPed.ColecaoPedJuros.Count = 0 Then
        Call mpInformarErro("Juros n�o localizado." & lMensagem, 2, True)
    End If

    For Each lobjstcPedJur In mobjstcPed.ColecaoPedJuros
        If lobjstcPedJur.Num_Parcela = mobjstcPed.Qtd_Parc Then
            If mobjstcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec = e_SimNao_Sim Then
                mobjstcPed.Per_Juro_Frac = lobjstcPedJur.Val_Ant
                mobjstcPed.objstcCtrlEmis.Per_Juros_Neg = lobjstcPedJur.Val_Ant
            Else
                mobjstcPed.Per_Juro_Frac = lobjstcPedJur.Val_Pos
                mobjstcPed.objstcCtrlEmis.Per_Juros_Neg = lobjstcPedJur.Val_Pos
            End If
            Exit For
        End If
    Next lobjstcPedJur


    'Compara��o do fracionamento
    If mobjstcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec = e_SimNao_Sim Then
        '018. Forma de pagamento: primeira parcela.
        mobjstcPed.Val_PriParcela = Format(lValorParc(mobjstcPed.Qtd_Parc, 1, 1), "standard")
        lDiferenca = mobjstcPed.Val_PriParcela - Format(gM46V999.gfTratarNulo(lrst("val_pr_antec"), tpoNumerico), "standard")
        If lDiferenca <= "0,02" And lDiferenca >= "-0,02" Then
            mobjstcPed.Val_PriParcela = Format(lrst("val_pr_antec"), "standard")
        Else
            If Format(mobjstcPed.Val_PriParcela, "standard") <> Format(gM46V999.gfTratarNulo(lrst("val_pr_antec"), tpoNumerico), "standard") Then
                If mobjstcPed.Cod_Agen_Bas_Cor <> 0 And mobjstcPed.objstcCtrlEmis.Num_Bloq = 0 Then
                    mobjstcPed.Val_PriParcela = 0
                Else
                    If gM46V999.gfTratarNulo(lrst("val_pr_antec"), tpoNumerico) <> 1 Then
                        Call mpInformarErro("Valor da 1� parcela divergente - Transmitido = R$" & Format(gM46V999.gfTratarNulo(lrst("val_pr_antec"), tpoNumerico), "standard") & ", calculado = R$" & Format(mobjstcPed.Val_PriParcela, "standard"), 2, True)
                    Else
                        mobjstcPed.Val_Pr_Antec = mobjstcPed.Val_PriParcela
                    End If
                End If
            End If
        End If
        If mobjstcPed.Cod_Agen_Bas_Cor <> 0 Then
            mobjstcPed.Val_PriParcela = 0
            mobjstcPed.Val_Pr_Antec = 0
        End If
    End If

    'Ficha 1000333
    frmT46V012B.flxResultado.Row = mLinha
    frmT46V012B.flxResultado.Col = 0
    mLinha = mLinha + 1
    frmT46V012B.flxResultado.AddItem "", mLinha
    mLinha = mLinha + 1
    frmT46V012B.flxResultado.AddItem "", mLinha
    mLinha = mLinha + 1
    frmT46V012B.flxResultado.AddItem "", mLinha
    frmT46V012B.flxResultado.TextMatrix(mLinha, 1) = "Quantidade de parcelas = " & Format(mobjstcPed.Qtd_Parc, "00")

    frmT46V012B.flxResultado.Row = mLinha
    frmT46V012B.flxResultado.Col = 0
    mLinha = mLinha + 1
    frmT46V012B.flxResultado.AddItem "", mLinha
    If mobjstcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec = e_SimNao_Sim Then
        frmT46V012B.flxResultado.TextMatrix(mLinha, 1) = "Valor das parcelas = R$ " & Format(lValorParc(mobjstcPed.Qtd_Parc, 1, 1), "standard")
    Else
        frmT46V012B.flxResultado.TextMatrix(mLinha, 1) = "Valor das parcelas = R$ " & Format(lValorParc(mobjstcPed.Qtd_Parc, 3, 1), "standard")
    End If

    If mQtdErros <> lQtdErro Then pDemaisErros = True

    Exit Sub

ControleErro:
    Call mpInformarErro(Err & " " & Err.Description, 2, True)

    'Obtem o inicio da mensagem de erro
    lMensagem = "mfCalPremios ERRO - " & Err & " " & Error & _
                Chr(13) & "gM46V999.CheckPoint " & lCkPoint & " - "

    Resume Next
    'simony
    'Carrega a variavel lCkStr
    On Error Resume Next
    lCkStr = " | " & ""

    On Error GoTo 0

    'Completa a mensagem de erro
    lMensagem = lMensagem & lCkStr

    Call mpInformarErro(lMensagem, 2, True)

End Sub

Private Function mfCalPreItem(pobjstcPed As stcA46V702B, pobjstcPedLoc As stcA46V708B, pobjstcItem As stcA46V708B) As Boolean
'Fun��o: Calcula pr�mios do item.
'Par�metros de entrada...
'                        pAtuStsItem.........Indica se o status do pr�mio pode ou n�o ser atualizado.
'Fun��o: calcula pr�mios.
'********************Vari�veis padr�o*************************************
    Dim lCkPoint As String    'Variavel utilizada para determinar pontos onde o c�digo j� passou.
    Dim lCkStr As String    'Variavel utilizada para guardar a posi��o das variaveis at� o ultimo
'lCkPoint. (posi��o das variaveis em ordem de declara��o)
'********************Vari�veis Declaradas para o processamento************
    Dim lAguarde As String    'Mensagem de aguarde.
    Dim lbdMulti As ADODB.Connection    'Conex�o do Banco local das TAB_PED tempor�rias
    Dim lChave As Long    'Posi��o na cole��o.
    Dim lErro As Variant    'Erros da DLL de c�lculo de pr�mio.
    Dim lI As Byte    'Utilizado com For...Next.
    Dim lMensagem As String    'Mensagem.
    Dim lMsgInsp As String    'Mensagem de inspe��o
    Dim lPrefixo As String    'Prefixo de mensagem.
    Dim lretorno As Integer    'Retorno.
    Dim lrst As clsYasRecordSet    'Registro.
    Dim lSelect As String    'Select
    Dim lUltData As Double    'Data da �ltima atualiza��o.
    Dim lobjstcCobert As stcA46V704B    'Objeto da classe (stc) de coberturas.
    Dim lobjCobert As clsA46V704A    'Objeto da classe de coberturas.
    Dim lobjPedObs As clsA46V714A    'Objeto da classe de observa��o.
    Dim lobjPedLoc As clsA46V708A    'Objeto da classe de itens.
    Dim lObjPed As clsA46V702A    'Objeto da classe de estrutura stcA46V702A (TAB_PED)
    Dim lBonus As Long
    Dim lColecaoTemp As Collection
    Dim AtualizaProgres As Boolean

    Dim lCod_Aceit As Integer

    mfCalPreItem = False

    'Calcula
    Set lobjPedLoc = New clsA46V708A
    Set lobjPedLoc.M46V999 = gM46V999

    Set lObjPed = New clsA46V702A
    Set lobjPedLoc.ConexaoBD = lObjPed.ConexaoBD
    Set lObjPed.M46V999 = gM46V999

    Set pobjstcItem.ObjPai = pobjstcPed

    Call mpConsisteItem(pobjstcPed, pobjstcPedLoc, pobjstcItem.Num_Item, lretorno)
    lCkPoint = "8"

    'Consiste item - n�vel 2.
    gM46V111.gCod_plano = mobjstcPed.Cod_Estip


    Call mpVerIteNiv_2(pobjstcPed, pobjstcItem)
    lCkPoint = "9"

    'd) Consiste item - normas
    'panAguarde.Caption = lAguarde & "Consistindo as normas..."
    'panAguarde.Refresh

    lCkPoint = "10"
    'Verifica Regras de Aceita��o Pontuais

    If gM46V999.VerificarNormas Then
        lCod_Aceit = 0
        lCod_Aceit = gM46V103.mfClasseAceit(pobjstcItem, pobjstcPed, "", "")

        If lCod_Aceit = 3 Then
            Call mpInformarErro("Proposta sem aceita��o.", 2, True)
        ElseIf lCod_Aceit = 4 Then
            Call mpInformarErro("Para essa proposta solicitar cota��o a Unidade de Produtos Diversos.", 2, True)
        End If

        lCkPoint = "11"
    End If

    Call mpVerNormas(pobjstcPed, pobjstcPedLoc, pobjstcItem.Num_Item, lretorno, lCod_Aceit, lMsgInsp)

    If lMsgInsp = "" And pobjstcItem.Ind_Sin_Inspec > 0 Then
        lMsgInsp = "OBRIGATORIO A REALIZACAO DE INSPECAO DO RISCO"
    End If

    If gM46V999.gfPreenchido(lMsgInsp) = True Then
        If pobjstcPed.Cod_Tip_Doc = e_TipDoc_Proposta Then
            If gM46V999.gfPreenchido(pobjstcItem.objstcPedInsp.Num_Tel) = False Then
                Call mpInformarErro("Telefone de contato para inspe��o inv�lido (n�o preenchido).", 2, True)
            End If

            If gM46V999.gfPreenchido(pobjstcItem.objstcPedInsp.Pess_Cont) = False Or Len(pobjstcItem.objstcPedInsp.Pess_Cont) <= 3 Then
                Call mpInformarErro("Contato para inspe��o inv�lido (n�o preenchido).", 2, True)
            End If
        End If

        lretorno = 0
    End If

    If Not gM46V999.gfAbrBasDados("P00MULTI2", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdMulti, lMensagem) Then
        Call mpInformarErro(lMensagem, 2, True)
        Call gM46V999.gpMenErro(lMensagem)
        Exit Function
    End If

    AtualizaProgres = True

    lSelect = "Select * from TAB_PED_LOC"
    lSelect = lSelect & " where num_item = " & pobjstcItem.Num_Item
    lSelect = lSelect & " and num_apol = 0"

    'Busca os pr�mios de coberturas
    If Not gM46V999.gfObtRegistro(lbdMulti, lSelect, lrst, lMensagem) Then
        Call mpInformarErro(lMensagem, 2, True)
        Call gM46V999.gpMenErro(lMensagem)
        Exit Function
    End If

    If Not lrst.EOF Then
        pobjstcItem.Per_Desc_Progres = lrst("Per_Desc_Progres")
    End If

    If Not lobjPedLoc.mfAlterar(pobjstcItem, "", "") Then
        lMensagem = lobjPedLoc.UltimoErro
        Exit Function
    End If

'    If gM46V999.gCodProduto <= gM46V999.TarifaAntiga Then
'        lSelect = "Select * from Tab_Ped_Cobert"
'        lSelect = lSelect & " where Tab_Ped_Cobert.num_item = " & pobjstcItem.Num_Item
'        lSelect = lSelect & " and Tab_Ped_Cobert.num_apol = 0"
'        lSelect = lSelect & " and Tab_Ped_Cobert.Val_IS <> 0"
'
'        'Busca os pr�mios de coberturas
'        If Not gM46V999.gfObtRegistro(lbdMulti, lSelect, lrst, lMensagem) Then
'            Call mpInformarErro(lMensagem, 2, True)
'            Call gM46V999.gpMenErro(lMensagem)
'            Exit Function
'        End If
'
'TentadeNovo:
'
'        Set pobjstcItem.ColecaoPedCobert = New Collection
'        Set lobjstcCobert = New stcA46V704B
'        Set lobjstcCobert.M46V999 = gM46V999
'
'        Set lobjCobert = New clsA46V704A
'        Set lobjCobert.M46V999 = gM46V999
'        'Set lObjCobert.ConexaoBD = objPedido.ConexaoBD
'
'        lobjstcCobert.NOSSO_NUMERO = pobjstcPed.NOSSO_NUMERO
'        lobjstcCobert.Num_Item = pobjstcItem.Num_Item
'
'        If Not lobjCobert.mfExcluir(lobjstcCobert, "", "", True) Then
'            Call gM46V999.gpMenErro(lobjCobert.UltimoErro)
'            Exit Function
'        End If
'
'        Do While Not lrst.EOF
'            Set lobjstcCobert = New stcA46V704B
'            Set lobjstcCobert.M46V999 = gM46V999
'
'            lobjstcCobert.Cod_Cobert = lrst("Cod_Cobert")
'            lobjstcCobert.Cod_Org = lrst("Cod_Org")
'            lobjstcCobert.Cod_Ramo = pobjstcPed.Cod_Ramo
'            lobjstcCobert.NOSSO_NUMERO = pobjstcPed.NOSSO_NUMERO
'            lobjstcCobert.Num_Item = pobjstcItem.Num_Item
'            lobjstcCobert.Num_Ped = pobjstcItem.Num_Ped
'            lobjstcCobert.Per_Desc_Coml = lrst("Per_Desc_Coml")
'            lobjstcCobert.Per_Desc_Exp = lrst("Per_Desc_Exp")
'            lobjstcCobert.Per_Franq = lrst("Per_Franq")
'            lobjstcCobert.Per_Taxa = lrst("Per_Taxa")
'            lobjstcCobert.Tab_Cobert = lrst("Tab_Cobert")
'            lobjstcCobert.Val_Franq_Max = lrst("Val_Franq_Max")
'            lobjstcCobert.Val_Franq_Min = lrst("Val_Franq_Min")
'            lobjstcCobert.Val_IS = lrst("Val_IS")
'            lobjstcCobert.VAL_PR = lrst("Val_Pr")
'            lobjstcCobert.Val_Pr_Anu = lrst("Val_Pr_Anu")
'            lobjstcCobert.Val_Pr_Resg = lrst("Val_Pr_Resg")
'            lobjstcCobert.Val_Pr_Resg_Anu = lrst("Val_Pr_Resg_Anu")
'            lobjstcCobert.Val_Pr_Tarif = lrst("Val_Pr_Tarif")
'            lobjstcCobert.Val_Pr_Tarif_Anu = lrst("Val_Pr_Tarif_Anu")
'
'            If Not lobjCobert.mfIncluir(lobjstcCobert, "", "", True, False, "") Then
'                Call gM46V999.gpMenErro(lobjCobert.UltimoErro)
'                Exit Function
'            End If
'
'            pobjstcItem.ColecaoPedCobert.Add lobjstcCobert, lobjstcCobert.Class_ID
'            lrst.MoveNext
'        Loop
'
'    End If

    lCkPoint = "50"

    'Verifica Risco em Shopping Center
    Dim lobjstcFator As stcA46V720B       'Objeto da classe de estrutura stcA46V720B (TAB_PED_FAT_RISCO).
    gM46V999.gAchouShopping = False
    If pobjstcPedLoc.Cod_Ramo = e_CodRamo_Empresarial Then
        For Each lobjstcFator In pobjstcPedLoc.ColecaoPedFator
            If lobjstcFator.Cod_Fator = 9 Then    'RISCO EM SHOPPING CENTER
                gM46V999.gAchouShopping = True
                Exit For
            End If
        Next lobjstcFator
    End If

    'Antes de Carregar objeto vai calcular novamente
    If gM46V999.gCodProduto > gM46V999.TarifaAntiga Then
        If gM46V111.gTipoGrupMulti = e_GrupoRamo_Multi Then
            pobjstcItem.Cod_Sit_item = e_SitItem_NaoCalc
            If Not gM46V103.mfCalNewPrecif(pobjstcItem, AtualizaProgres, lMensagem, "", "", "") Then
                Call mpInformarErro(lMensagem, 2, True)
                Call gM46V999.gpMenErro(lMensagem)
                Exit Function
            End If
        End If
    End If

    gCarga = True
    'txtDescProgr.Text = pobjstcItem.Per_Desc_Progres
    gCarga = False

    'Salva resultado do c�lculo.
    'a) Atualiza registro do item.
    lUltData = CDbl(gM46V999.gfForData(3))
    pobjstcItem.Cod_Sit_item = e_SitItem_Calculado

    'If pobjstcItem.Num_Item = Val(txtNumItem.Text) Then
    '    gIteUltUsuario = gM46V999.gNomUsuario6
    '    gIteUltData = lUltData
    '    Call mpCarregarCobertGrid(pobjstcItem.ColecaoPedCobert)
    'End If

    pobjstcPed.ColecaoPedLoc.Remove pobjstcItem.Class_ID


    'Verifico se deve ser gravada a observa��o de inspe��o.
    Dim lTexto As String
    
    lTexto = frmT46V101A.Nova_Observacao(pobjstcItem.Ind_StatusInspecao, pobjstcItem.Num_Insp, pobjstcItem)
    If gM46V999.gfPreenchido(lMsgInsp) And pobjstcItem.Ind_StatusInspecao = e_StatusInspecao_ConfNormas Then
        If gM46V999.gNovaPropostaLMI Then
            lTexto = lTexto & "Item " & Format(pobjstcItem.Num_Item, "00000") & ": " & gM46V999.gfRetirarCaracterEspecial(lMsgInsp, True)
        Else
            lTexto = lTexto & gM46V999.gfRetirarCaracterEspecial(lMsgInsp, True)
        End If
    End If
    
    UCase (lTexto)
    
    Set lobjPedObs = New clsA46V714A
    Set lobjPedObs.M46V999 = gM46V999
    
    Set lobjPedObs.ConexaoBD = lObjPed.ConexaoBD
    If Not lobjPedObs.mfExcluir(pobjstcItem.objstcPedObs, "", "") Then
        lMensagem = lobjPedObs.UltimoErro
        Call gM46V999.gpMenErro(lMensagem)
        Exit Function
    End If
    
    If Trim(lTexto) <> "" Then
        With pobjstcItem.objstcPedObs
            .NOSSO_NUMERO = pobjstcItem.NOSSO_NUMERO
            .Num_Item = pobjstcItem.Num_Item
            .Observacao = Replace(lTexto, vbCrLf, "&%")
        End With
        If Not lobjPedObs.mfIncluir(pobjstcItem.objstcPedObs, "", "", "") Then
            lMensagem = lobjPedObs.UltimoErro
            Call gM46V999.gpMenErro(lMensagem)
            Exit Function
        End If
    End If

    pobjstcPed.ColecaoPedLoc.Add pobjstcItem, pobjstcItem.Class_ID
    mfCalPreItem = True

End Function

Public Sub mpObtemMaxParcelas(pobjstcPed As stcA46V702B, _
                              pobjstcPedLoc As stcA46V708B, _
                              ByVal pMaxParcPerm As Integer, _
                              ByRef lParcsVista As Integer, ByRef lParcs30dia As Integer, _
                              Optional lUltParcela As Double)
'Determina quantide m�xima de parcelas
'Parametros de entrada/saida      lParcsVista ... parcelas com pagamento antecipado
'                                 lParcs30dia ... parcelas com pgto Postecipado

    Dim lAchouNParc As Boolean
    Dim lAuxData As Long
    Dim lTotalDIAS As Integer
    Dim lDataVen As Variant    'Retorna todas as datas.
    Dim lDataUltParc As Double    'Retorna data da ultima parcela possivel
    Dim lIrregular As Boolean    'Indica se h� data irregula.
    Dim lHoje As Long    'Dia de hj.
    Dim lDiasAnteTerVig As Integer      'Numero de dias antes do t�rmino de vigencia
    'para determinar o vencimento da ultima parcela
    Dim lretorno As String    'Retorno da DLL.
    Dim lMensagem As String    'Mensagem.
    Dim lDataFinal As Double

    '1 inicia o valor das variaveis
    lParcs30dia = 0
    lParcsVista = 0
    If pMaxParcPerm = 0 Then pMaxParcPerm = 12

    '2- Determina o numero m�ximo de dias antes do t�rmino de vigencia para a
    'ultima parcela.
    lDiasAnteTerVig = 30
    lHoje = Format(Date, "yyyymmdd")

    '3- Obtem o numero de parcelas a vista

    lDataFinal = mfCalcData(pobjstcPed.Dat_Ter_Vig, (lDiasAnteTerVig * -1))

    If Right(lDataFinal, 4) = "0229" Then
        lDataFinal = Left(lDataFinal, 4) & "0228"
    End If

    If pobjstcPed.Dat_Ini_Vig <> 0 And pobjstcPed.Dat_Ter_Vig <> 0 Then
        If lHoje < pobjstcPed.Dat_Ini_Vig Then
            'Inicio de vigencia � a referencia p/ a qtd de parcelas
            lTotalDIAS = val(gM46V999.gfDifDatas(pobjstcPed.Dat_Ini_Vig, lDataFinal)) - 10
        Else
            'Inicio de vigencia menor que a data atual
            'Data atual � referencia p/ a qtd de parcelas
            lTotalDIAS = val(gM46V999.gfDifDatas(lHoje, lDataFinal)) - 10
        End If
    End If

    If lTotalDIAS > 30 Then
        lParcsVista = lTotalDIAS / 30
    Else
        If lTotalDIAS > 0 Then
            lParcsVista = 1
        Else
            lParcsVista = 0
        End If
    End If

    If lParcsVista > 0 Then
        lAchouNParc = False
        lParcsVista = Int(lParcsVista) + 1
    Else
        ' Vai retornar 0
        lAchouNParc = True
    End If

    '4- Obtem maximo de parcelas com a 1� em 30 Dias
    lDataFinal = mfCalcData(pobjstcPed.Dat_Ter_Vig, (lDiasAnteTerVig * -1))

    If Right(lDataFinal, 4) = "0229" Then
        lDataFinal = Left(lDataFinal, 4) & "0228"
    End If

    If pobjstcPed.Dat_Ini_Vig <> 0 And pobjstcPed.Dat_Ter_Vig <> 0 Then
        If lHoje < pobjstcPed.Dat_Ini_Vig Then
            'Inicio de vigencia � a referencia p/ a qtd de parcelas
            lTotalDIAS = val(gM46V999.gfDifDatas(pobjstcPed.Dat_Ini_Vig, lDataFinal)) - 10
        Else
            'Inicio de vigencia menor que a data atual
            'Data atual � referencia p/ a qtd de parcelas
            lTotalDIAS = val(gM46V999.gfDifDatas(lHoje, lDataFinal)) - 10
        End If
    End If

    If lTotalDIAS > 30 Then
        lParcs30dia = lTotalDIAS / 30
    Else
        If lTotalDIAS > 0 Then
            lParcs30dia = 1
        Else
            lParcs30dia = 0
        End If
    End If

    If lParcs30dia > 0 Then
        lAchouNParc = False
        lParcs30dia = Int(lParcs30dia)
    Else
        ' Vai retornar 0
        lAchouNParc = True
    End If

    '5- Limita as quantidades de parcela ao limite m�ximo permitido
    If lParcsVista > pMaxParcPerm Then
        lParcsVista = pMaxParcPerm
    End If

    If lParcs30dia > pMaxParcPerm Then
        lParcs30dia = pMaxParcPerm
    End If

    '6- Ajusta parcela de endosso para pelo menos 1
    If pobjstcPedLoc.Tip_Emissao >= 300 And lParcsVista = 0 Then
        If pobjstcPed.Dat_Ter_Vig <> 0 Then
            If val(gM46V999.gfDifDatas(lHoje, pobjstcPed.Dat_Ter_Vig)) >= 0 Then
                lParcsVista = 1
            End If
        End If
    End If

    If (pobjstcPedLoc.Tip_Emissao = 301 Or pobjstcPedLoc.Tip_Emissao = 303) And lParcs30dia = 0 Then
        lParcs30dia = 1
    End If

    '7- Ajusta numero de parcelas para perfil 1
    If gM46V999.PerfilAcesso = e_Perfil_DIG_Gestor Then
        'lParcsVista = pMaxParcPerm         'Comentada devida a ficha 0600403
        'lParcs30dia = pMaxParcPerm         'Comentada devida a ficha 0600403
        'Devida a ficha n� 0600403 dever� ser consistido para o usu�rio digita��o gestor
        'que o n�mero de parcelas dever� ser a quantidade permitida e n�o a m�xima do produto
        If lParcsVista = 0 Then
            lParcsVista = 1
        End If

        If lParcs30dia = 0 Then
            lParcs30dia = 1
        End If
    End If

    '8- Grava na variavel opcional de retorno a data limite da ultima parcela
    If pobjstcPed.Dat_Ter_Vig <> 0 Then
        lUltParcela = lDataFinal
    End If

End Sub
Private Function mfCalcData(pData As Long, pDias As Integer) As Long
    Dim lAuxPreData As Date
    If pData <> 0 Then

        If Right(pData, 4) = "0229" Then
            pData = Left(pData, 4) & "0228"
        End If

        If Mid(pData, 7, 2) & "/" & Mid(pData, 5, 2) & "/" & Mid(pData, 1, 4) <> Format(Mid(pData, 7, 2) & "/" & Mid(pData, 5, 2) & "/" & Mid(pData, 1, 4), "dd/mm/yyyy") Then
            lAuxPreData = CDate(Mid(pData, 5, 2) & "/" & Mid(pData, 7, 2) & "/" & Mid(pData, 1, 4))
        Else
            lAuxPreData = CDate(Mid(pData, 7, 2) & "/" & Mid(pData, 5, 2) & "/" & Mid(pData, 1, 4))
        End If

        mfCalcData = Format(CDate(lAuxPreData) + pDias, "yyyymmdd")
    End If

End Function

Private Sub mpPesquisaDigitacao()
    Dim lMensagem As String     'Mensagem
    Dim lNossoNumero As String        'Nosso Numero.
    Dim lNovNosNumero As String         'Novo: Nosso Numero.
    Dim lNovNumPedido As String         'Novo: n�mero do pediro.
    Dim lControl As Integer
    Dim lobjCtrlEmis As clsA46V712A   'Objeto da classe clsA46V712A (TAB_CTRL_EMIS)
    Dim lobjstcCtrlEmis As stcA46V712B      'Objeto da classe de estrutura stcA46V712B (TAB_CTRL_EMIS)
    Dim bdBasDados As ADODB.Connection    'Base de dados.
    Dim lObjPed As clsA46V702A    'Objeto da classe clsA46V702A (TAB_PED).
    Dim lObjstcPed As stcA46V702B    'Objeto da classe de estrutura stcA46V702B (TAB_PED).
    Dim lretorno As Integer
    Dim UltimoErro As String

    If gM46V999.gfAbrBasDados("P00MPED", gM46V999.gTipP00MPED, gM46V999.gEndP00MPED, bdBasDados, _
                              lMensagem) = False Then
        Call gM46V999.gpGraLog(1, "FNP0044 - " & lMensagem)
        Exit Sub
    End If

    If Not mfConsistirDados Then Exit Sub

    lControl = 0
    
Retorno:
    'Tenta localizar os dados no tab_ctrl_emis pelo PI.

    'Instancio os objetos
    Set lobjCtrlEmis = New clsA46V712A
    Set lobjCtrlEmis.M46V999 = gM46V999

    Set lobjstcCtrlEmis = New stcA46V712B
    Set lobjCtrlEmis.ConexaoBD = bdBasDados

    Set lObjPed = New clsA46V702A
    Set lObjPed.M46V999 = gM46V999

    Set lObjstcPed = New stcA46V702B
    Set lObjstcPed.M46V999 = gM46V999

    If gM46V999.gfPreenchido(txtNumeroPI) Then
        lobjstcCtrlEmis.NUM_PI = txtNumeroPI
        If Not lobjCtrlEmis.mfCarregarColecao(lobjstcCtrlEmis, "", "") Then
            UltimoErro = lobjCtrlEmis.UltimoErro
            Call gM46V999.gpGraLog(1, "FNP0024 - " & UltimoErro)
            Exit Sub
        End If
        If lobjCtrlEmis.Colecao.Count <> 0 Then
            '*****Localizou o PI
            Set lobjstcCtrlEmis = lobjCtrlEmis.Colecao.Item(1)
            txtPropostaSyas = lobjstcCtrlEmis.NOSSO_NUMERO
            If lobjstcCtrlEmis.Dat_Rec_Cia < DateAdd("d", -15, Date) And gM46V999.PerfilAcesso <> e_Perfil_DIG_Gestor And lobjstcCtrlEmis.Num_Apol = 0 Then
                Call gM46V999.gpGraLog(1, "FGE0079 - Data de recebimento inv�lida (inferior a 15 dias da data de " & Date & ")")
                txtDataRecebimento.Text = lobjstcCtrlEmis.Dat_Rec_Cia
                If txtDataRecebimento.Visible = True And txtDataRecebimento.Enabled = True Then
                    txtDataRecebimento.SetFocus
                End If
                Exit Sub
            End If
            lobjstcCtrlEmis.NUM_PI = txtNumeroPI
            lobjstcCtrlEmis.Dat_Rec_Cia = gM46V999.gfForData(1, frmT46V012A.txtDataRecebimento)
            If gM46V999.gfPreenchido(txtMTVVD.Text) And cboOpcao.ItemData(cboOpcao.ListIndex) = 3 Then
                lobjstcCtrlEmis.Num_Ref_Corr = "MT" & txtMTVVD.Text
            End If
            'Ficha 1000449 - Incluir campo Referencia Marsh para digita��o
            If gM46V999.gfPreenchido(txtMTVVD.Text) And cboOpcao.ItemData(cboOpcao.ListIndex) = 4 Then
                lobjstcCtrlEmis.Num_Ref_Corr = txtMTVVD.Text
            End If
    
            If Not lobjCtrlEmis.mfAlterarNum_PI(lobjstcCtrlEmis, "", "") Then
                UltimoErro = lobjCtrlEmis.UltimoErro
                Call gM46V999.gpGraLog(1, "FNS0024 - " & UltimoErro)
                Exit Sub
            End If
            lObjstcPed.NOSSO_NUMERO = lobjstcCtrlEmis.NOSSO_NUMERO
            If Not lObjPed.mfCarregarColecao(lObjstcPed, "", "") Then
                UltimoErro = lObjPed.UltimoErro
                Call gM46V999.gpGraLog(1, "FNP0024 - " & UltimoErro)
                Exit Sub
            End If
            If lObjPed.Colecao.Count <> 0 Then
                Set lObjstcPed = lObjPed.Colecao.Item(1)
                lObjstcPed.NUM_PI = txtNumeroPI
                If Not lObjPed.mfAlterar(lObjstcPed, "", "") Then
                    UltimoErro = lObjPed.UltimoErro
                    Call gM46V999.gpGraLog(1, "FNP0025 - " & UltimoErro)
                    Exit Sub
                End If
            End If
            'Ficha 1000902
            If cboOpcao.ListIndex <> 4 Then
                If gM46V111.gfGraRegComunicacao(lobjstcCtrlEmis.NOSSO_NUMERO, lMensagem) = False Then
                    Call gM46V999.gpGraLog(2, lMensagem)
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda And gM46V999.gTipoAcesso = DIGITACAO Then
                    frmT46V190A.NossoNumeroOld = lobjstcCtrlEmis.NOSSO_NUMERO
                    TemOutroFormAberto = True
                    Call frmT46V190A.Show(1)
                    TemOutroFormAberto = False
                    If frmT46V190A.TodosItensCorretos = False Then
                        Call gM46V999.gpGraLog(1, "Todos os itens que tem inspe��o obrigat�ria devem possuir dados de contato.")
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                    Call gM46V999.updAtualizaStatusInspecao(txtNumeroPI.Text, lMensagem, lObjstcPed)
                End If
            Else
                Call Shell(gM46V999.gEndExeP47V001 & " \" & Format(frmT46V012A.txtNumeroPI.Text, String(10, "0")) & "\" & Format(CDate(frmT46V012A.txtDataRecebimento), "yyyymmdd"), vbNormalFocus)
                GoTo Pular
            End If
            Me.Hide
            Exit Sub
        End If
    End If 'gM46V999.gfPreenchido(txtNumeroPI) Then
    
    If gM46V999.gfPreenchido(txtPropostaSyas.Text) Then
        If Not gM46V999.gBaseLocal Then
            lNossoNumero = txtPropostaSyas.Text
            If Not mpVerificaUnicidade(bdBasDados, lNossoNumero, gM46V999.CDblx(txtDV.Text), lretorno, lMensagem) Then
                Call gM46V999.gpGraLog(2, lMensagem)
                If txtDV.Visible = True And txtDV.Enabled = True Then
                    txtDV.SetFocus
                End If
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            
            
            If lretorno = 1 Then
                '****Transmiss�o n�o localizada******
                lobjstcCtrlEmis.NUM_PI = 0
                lobjstcCtrlEmis.NOSSO_NUMERO = txtPropostaSyas.Text
                If Not lobjCtrlEmis.mfCarregarColecao(lobjstcCtrlEmis, "", "") Then
                    UltimoErro = lobjCtrlEmis.UltimoErro
                    Call gM46V999.gpGraLog(1, "FNP0024 - " & UltimoErro)
                    Exit Sub
                End If
                If lobjCtrlEmis.Colecao.Count <> 0 Then
                    Set lobjstcCtrlEmis = lobjCtrlEmis.Colecao.Item(1)
                    If Not mfGeraCopia(lobjstcCtrlEmis.NOSSO_NUMERO, txtNumeroPI, gM46V999.gfForData(1, txtDataRecebimento), lNovNosNumero, lMensagem) Then
                        Call gM46V999.gpGraLog(2, lMensagem)
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    Else
                        If gM46V111.gfGraRegComunicacao(lNovNosNumero, lMensagem) = False Then
                            Call gM46V999.gpGraLog(2, lMensagem)
                            Screen.MousePointer = vbDefault
                            Exit Sub
                        End If
                        If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda And gM46V999.gTipoAcesso = DIGITACAO Then
                            frmT46V190A.NossoNumeroOld = txtPropostaSyas.Text
                            TemOutroFormAberto = True
                            Call frmT46V190A.Show(1)
                            TemOutroFormAberto = False
                            If frmT46V190A.TodosItensCorretos = False Then
                                Call gM46V999.gpGraLog(1, "Todos os itens que tem inspe��o obrigat�ria devem possuir dados de contato.")
                                Screen.MousePointer = vbDefault
                                Exit Sub
                            End If
                            Call gM46V999.updAtualizaStatusInspecao(txtNumeroPI.Text, lMensagem, lObjstcPed)
                        End If
                        Me.Hide
                        Exit Sub
                    End If
                End If
                Call gM46V999.gpGraLog(1, "FNL0041 - N�mero da proposta n�o localizado (" & txtPropostaSyas.Text & ").")
            ElseIf lretorno = 2 Then
                '****Transmiss�o localizada com DV diferente******
                Call gM46V999.gpGraLog(1, "FNL0041 - Proposta localizada com DV diferente do informado.")
            Else
                '****Transmiss�o localizada******
                'Calcula nosso n�mero e n�mero do pedido.
                If gM46V999.gfCalNosNumero(lNovNosNumero, lMensagem) = False Then
                    Call gM46V999.gpGraLog(1, "FNP0021 - Pesquisa Digita��o: " & lMensagem)
                    Exit Sub
                End If

                If gM46V999.gfCalNumPedido(lNovNumPedido, lMensagem) = False Then
                    Call gM46V999.gpGraLog(1, "FNP0022 - Pesquisa Digita��o: " & lMensagem)
                    Exit Sub
                End If

                Dim lSql As String
                Dim lrsTabela As clsYasRecordSet    'Registro da tabela.
                
                lSql = "SELECT COD_ESTIP FROM FTP_Tab_Ped WHERE Nosso_Numero = '" & lNossoNumero & "'"
                
                If gM46V999.gfObtRegistro(bdBasDados, lSql, lrsTabela, lMensagem) = False Then
                    Call gM46V999.gpGraLog(1, lMensagem)
                    Exit Sub
                End If
                
                If lrsTabela!Cod_Estip = 20 Then
                
                    If Not ReentradaImobiliario(txtNumeroPI.Text, lNossoNumero, lNovNosNumero, lMensagem) Then
                        Call gM46V999.gpGraLog(1, lMensagem)
                        Exit Sub
                    End If
                    
                Else
                        If Not gM46V001.gfImportarTransmissao(lNossoNumero, gM46V999.CDblx(txtDV.Text), lNovNosNumero, lNovNumPedido, txtNumeroPI, gM46V999.gfForData(1, frmT46V012A.txtDataRecebimento), lMensagem) Then
                            UltimoErro = lMensagem
                            Call gM46V999.gpGraLog(1, "FNP0024 - " & UltimoErro)
                            Exit Sub
                        End If
                End If

                lobjstcCtrlEmis.NOSSO_NUMERO = lNovNosNumero

                If Not lobjCtrlEmis.mfCarregarColecao(lobjstcCtrlEmis, "", "") Then
                    UltimoErro = lobjCtrlEmis.UltimoErro
                    Call gM46V999.gpGraLog(1, "FNP0024 - " & UltimoErro)
                    Exit Sub
                End If

                If lobjCtrlEmis.Colecao.Count <> 0 Then

                    Set lobjstcCtrlEmis = lobjCtrlEmis.Colecao.Item(1)
                    txtPropostaSyas = lobjstcCtrlEmis.NOSSO_NUMERO
                    lobjstcCtrlEmis.NUM_PI = txtNumeroPI
                    lobjstcCtrlEmis.Dat_Rec_Cia = gM46V999.gfForData(1, frmT46V012A.txtDataRecebimento)
                    If Not lobjCtrlEmis.mfAlterarNum_PI(lobjstcCtrlEmis, "", "") Then
                        UltimoErro = lobjCtrlEmis.UltimoErro
                        Call gM46V999.gpGraLog(1, "FNS0024 - " & UltimoErro)
                        Exit Sub
                    End If

                    If gM46V111.gfGraRegComunicacao(lobjstcCtrlEmis.NOSSO_NUMERO, lMensagem) = False Then
                        Call gM46V999.gpGraLog(2, lMensagem)
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                    If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda And gM46V999.gTipoAcesso = DIGITACAO Then
                        frmT46V190A.NossoNumeroOld = txtPropostaSyas.Text
                        TemOutroFormAberto = True
                        Call frmT46V190A.Show(1)
                        TemOutroFormAberto = False
                        If frmT46V190A.TodosItensCorretos = False Then
                            Call gM46V999.gpGraLog(1, "Todos os itens que tem inspe��o obrigat�ria devem possuir dados de contato.")
                            Screen.MousePointer = vbDefault
                            Exit Sub
                        End If
                        Call gM46V999.updAtualizaStatusInspecao(txtNumeroPI.Text, lMensagem, lObjstcPed)
                    End If
                    Me.Hide
                    Exit Sub
                End If
            End If
        End If
    End If 'gM46V999.gfPreenchido(txtPropostaSyas.Text) Then

    If CDate(frmT46V012A.txtDataRecebimento) < DateAdd("d", -15, Date) And gM46V999.PerfilAcesso <> e_Perfil_DIG_Gestor Then
        Call gM46V999.gpGraLog(1, "FGE0080 - Data de recebimento inv�lida (inferior a 15 dias da data de " & Date & ")")
        If txtDataRecebimento.Visible = True And txtDataRecebimento.Enabled = True Then
            txtDataRecebimento.SetFocus
        End If
        Exit Sub
    End If

    Screen.MousePointer = vbDefault
    gM46V111.gNum_Pi = txtNumeroPI
    TemOutroFormAberto = True
    frmT46V024A.Show 1
    TemOutroFormAberto = False

Pular:

    gM46V111.gNum_Pi = 0
    If gM46V999.gFormCancelado = False Then
        Me.Hide
    End If
End Sub

Private Sub mpValidaParcela(pobjstcPed As stcA46V702B, pobjstcPedLoc As stcA46V708B)
    Dim lMaxParcela As Integer  'N�mero m�ximo de parcelas permitido
    Dim lMaxParcelaAux As Integer     'N�mero m�ximo de parcelas permitido
    Dim lMaxParcelaVista As Integer       'N�mero m�ximo de parcelas permitido a vista
    Dim lMaxParcela30Dia As Integer       'N�mero m�ximo de parcelas permitido 30 dias
    Dim lDataUltVencimento As Double          'Data do ultimo vencimento possivel
    Dim lQutdeParcelas As Integer     'Quantidade de parcelas
    Dim lobjstcParc As stcA46V713B

    'Determina quantide m�xima de parcelas
    Call mpObtemMaxParcelas(pobjstcPed, pobjstcPedLoc, gM46V999.gProQuaMaxParcelas, lMaxParcelaVista, lMaxParcela30Dia, lDataUltVencimento)

    'Informa o maior valor de parcelas
    If lMaxParcelaVista > lMaxParcela30Dia Then
        lMaxParcela = lMaxParcelaVista
    Else
        lMaxParcela = lMaxParcela30Dia
    End If

    'Guarda na variavel o numero de parcelas informado
    lQutdeParcelas = pobjstcPed.Qtd_Parc

    'caso seja proposta limita o numero maximo para o tipo
    'escolhido (antecipado ou 30dias)
    If pobjstcPed.Cod_Tip_Doc = e_TipDoc_Proposta Then
        If pobjstcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec = e_SimNao_Sim Then
            lMaxParcela = lMaxParcelaVista
        Else
            lMaxParcela = lMaxParcela30Dia
        End If

        'ENDOSSO
        'Ap�s o c�lculo � que ser� consistido se a postecipada � permitida ou n�o.
        'atrav�s da vari�vel lMaxParcelaAux.
        lMaxParcelaAux = lMaxParcela

        If lMaxParcela = 0 And pobjstcPedLoc.Tip_Emissao >= 300 Then
            lMaxParcela = 1
        End If
    End If

    'Plano Mercedes e RVA
    If (gM46V111.gCod_plano >= 2021 And gM46V111.gCod_plano <= 2031) Or _
       (gM46V111.gCod_plano >= 2326 And gM46V111.gCod_plano <= 2349) Or _
       (gM46V111.gCod_plano >= 2350 And gM46V111.gCod_plano <= 2373) Or _
       (gM46V111.gCod_plano = 2303 And pobjstcPed.Cod_Ramo = e_CodRamo_Empresarial) Or _
       (gM46V111.gCod_plano = 2374 And pobjstcPed.Cod_Ramo = e_CodRamo_Residencial) Then
        lMaxParcela = 11
    End If

    If pobjstcPed.Cod_Tip_Doc = e_TipDoc_Cotacao Then
        'ajusta o numero de parcelas caso retorne 0
        If lMaxParcela <= 0 Then
            lMaxParcela = 1
        End If

        'Se = Cota��o ent�o atualiza o campo
        If pobjstcPed.Qtd_Parc > lMaxParcela Then
            pobjstcPed.Qtd_Parc = lMaxParcela
        End If
    Else
        'Verifica se o numero informado � maior que o permitido
        If lQutdeParcelas > lMaxParcela Then
            If pobjstcPed.ColecaoPedVenc.Count > 0 And gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then
                For Each lobjstcParc In pobjstcPed.ColecaoPedVenc
                    If lobjstcParc.DAT_VENC > pobjstcPed.Dat_Ter_Vig Then
                        Call mpInformarErro("Quantidade de parcelas inv�lida (Data da �ltima parcela n�o deve ultrapassar " & _
                                            gM46V999.gfForData(2, lobjstcParc.DAT_VENC) & ". " & val(gM46V999.gfDifDatas(lDataUltVencimento, pobjstcPed.Dat_Ter_Vig)) & _
                                          " dias antes do t�mino da vig�ncia  )." & vbCrLf & "N�mero m�ximo de parcelas permitido: " & lMaxParcela, 2, True)
                        Exit Sub
                    End If

                Next lobjstcParc
                lMaxParcela = lQutdeParcelas
            Else
                If lDataUltVencimento = 0 Then
                    Call mpInformarErro("Quantidade de parcelas inv�lida (Data da �ltima parcela deve ser inferior a data de t�rmino da vig�ncia)." & vbCrLf & "N�mero m�ximo de parcelas permitido: " & lMaxParcela, 2, True)
                Else
                    If val(gM46V999.gfDifDatas(lDataUltVencimento, pobjstcPed.Dat_Ter_Vig)) > 0 Then
                        Call mpInformarErro("Quantidade de parcelas inv�lida (Data da �ltima parcela n�o deve ultrapassar " & _
                                            gM46V999.gfForData(2, lDataUltVencimento) & ". " & val(gM46V999.gfDifDatas(lDataUltVencimento, pobjstcPed.Dat_Ter_Vig)) & _
                                          " dias antes do t�mino da vig�ncia  )." & vbCrLf & "N�mero m�ximo de parcelas permitido: " & lMaxParcela, 2, True)
                    Else
                        Call mpInformarErro("Quantidade de parcelas inv�lida (Data da �ltima parcela deve ser inferior a data de t�rmino da vig�ncia)." & vbCrLf & "N�mero m�ximo de parcelas permitido: " & lMaxParcela, 2, True)
                    End If
                End If
                Exit Sub
            End If
        End If
    End If

End Sub


Private Sub mpVerComum(pobjstcPed As stcA46V702B, pobjstcPedLoc As stcA46V708B)
'Procedure: consiste dados comuns

    Dim lTipPessoa As Byte    'Tipo da pessoa: 1 = pessoa jur�dica e 2 = pessoa f�sica.
    Dim lI As Integer    'Utilizado no for...next
    Dim lMaxParcela As Integer  'N�mero de parcelas m�ximo permitido.
    Dim clsA46V102 As clsA46V102A    'DLL : consiste dados comuns (n�vel 2).
    Dim lObjPed As clsA46V702A    'Objeto da classe de estrutura stcA46V702A (TAB_PED)
    Dim lObjstcVenc As stcA46V713B    'Objeto da classe de estrutura stcA46V713B (TAB_PED_Venc)
    Dim lObjstcCorr As stcA46V705B    'Objeto da classe de estrutura stcA46V705B (TAB_PED_Corr)
    Dim lobjstcCoss As stcA46V706B    'Objeto da classe de estrutura stcA46V706B (TAB_PED_Coss)
    Dim lCosseguros(5, 2) As String         'Cosseguros.
    Dim lCod_Retorno As Integer
    Dim lMensagem As String
    Dim lDigito As String
    Dim lAuxiliar As String
    Dim lPercAux As Single
    Dim lCodCorr As Long
    Dim lBlnSiscad As Boolean    'Indicador se existe dados no Siscad
    Dim objTabVenc As clsA46V713A
    Dim lArrayDatVenc As Variant
    Dim lDataIrregular As Boolean
    Dim txtobs As String
    Dim lClassif As String

    'lClassif = Val(Trim(Left(frmT46V101A.txtClassificacao.Text, 3)) + Trim(Mid(frmT46V101A.txtClassificacao.Text, 6, 2)))

    lClassif = pobjstcPedLoc.Cod_Clasf & pobjstcPedLoc.Cod_Compl_Clasf

    gM46V999.gQuaErros = 0
    lPercAux = 0
    lCodCorr = 0

    If pobjstcPed.ColecaoPedCorr.Count > 0 Then
        Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(1)
        lPercAux = lObjstcCorr.Per_Corr
        lCodCorr = lObjstcCorr.Cod_Corr
    End If

    pobjstcPed.objstcCtrlEmis.Cod_Envi_Doc = "X"
    If gM46V999.gfVerEtiqueta(lCodCorr, pobjstcPed.Cod_Reg, pobjstcPed.Tip_Pessoa, pobjstcPed.objstcCtrlEmis.Cod_Envi_Doc, lMensagem) Then
        If lMensagem = "Envio do documento deve ser para o 'Corretor'." Then
            pobjstcPed.objstcCtrlEmis.Cod_Envi_Doc = "C"
        Else
            pobjstcPed.objstcCtrlEmis.Cod_Envi_Doc = "S"
        End If
    Else
        '[FRH]Ficha:0801295
        'Novas op��es "Tipo de Emiss�o=Conforme cadastro":
        '"O" -> Conforme Cadastro Corretor e "E" -> Conforme Cadastro Segurado
        'a defini��o se "O" ou "E" � feita dentro da gM46V999.gfVerEtiqueta
        If lMensagem = "O" Then    '<-Triagem
            pobjstcPed.objstcCtrlEmis.Cod_Envi_Doc = "O"
        ElseIf lMensagem = "E" Then
            pobjstcPed.objstcCtrlEmis.Cod_Envi_Doc = "E"
        End If
    End If

    If mobjstcPed.objstcCtrlEmis.Tip_Cotac = 0 Then
        mobjstcPed.objstcCtrlEmis.Tip_Cotac = 1
    End If

    If mobjstcPed.Ind_CotacaoSolic = True And _
       mobjstcPed.StatusCotSiscota = enum_StatusTransm_Aprovada Then
        gM46O001.PermiteDigitarSemNormas = True
        gM46O001.PermiteQQDescontoComercial = True
        gM46V999.VerificarNormas = False
    End If

    '1. Consiste dados comuns - n�vel 1 (dados para c�lculo).
    Call mpConsisteComum(pobjstcPed, pobjstcPedLoc, lCod_Retorno)         'DLL

    If Right(pobjstcPed.Dat_Ini_Vig, 4) = "0229" Then
        Call mpInformarErro("Vig�ncia inv�lida (data de in�cio com ano " & _
                            "bissexto - utilizar 28/02 ou 01/03).", 2, True)
    End If

    '2. Valida as parcelas
    Call mpValidaParcela(pobjstcPed, pobjstcPedLoc)

    If pobjstcPed.Cod_Tip_Doc = e_TipDoc_Proposta Then

        '2. Consiste dados comuns - n�vel 2.
        Set clsA46V102 = New clsA46V102A    'DLL : consiste dados comuns (n�vel 2).
        Set clsA46V102.M46V999 = gM46V999

        With clsA46V102
            'a) Prepara dados para consist�ncia de comum n�vel 2.
            .Pro_CarBanco = pobjstcPed.Cod_Bco_Cobr
            'Obs.: O campo Ven_Parcela1_VM cont�m a data de vencimento da primeira parcela do carn�.

            If pobjstcPed.ColecaoPedVenc.Count > 0 Then
                Set lObjstcVenc = pobjstcPed.ColecaoPedVenc.Item(1)
                .Pro_CarDatPriParcela = lObjstcVenc.DAT_VENC
            End If

            .Pro_Cob_Bairro = pobjstcPed.Nom_Bairro
            .Pro_Cob_CEP = Format(pobjstcPed.Num_CEP, "00000000")
            .Pro_Cob_Cidade = pobjstcPed.Nom_Cid

            If gM46V999.gfPreenchido(pobjstcPed.Dsc_Compl_Logr) Then
                .Pro_Cob_Complemento = pobjstcPed.Dsc_Compl_Logr
            Else
                .Pro_Cob_Complemento = ""
            End If

            .Pro_Cob_NomLogradouro = pobjstcPed.Nom_Logr
            .Pro_Cob_TipLogradouro = pobjstcPed.Cod_Tip_Logr
            .Pro_Cob_UF = pobjstcPed.Sig_UF

            If pobjstcPed.ColecaoPedCorr.Count > 0 Then
                Set lObjstcCorr = pobjstcPed.ColecaoPedCorr.Item(1)
                .Pro_CodCorretor = lObjstcCorr.Cod_Corr
            Else
                .Pro_CodCorretor = 0
            End If

            lCosseguros(1, 1) = 0
            lCosseguros(1, 2) = 0
            lCosseguros(2, 1) = 0
            lCosseguros(2, 2) = 0
            lCosseguros(3, 1) = 0
            lCosseguros(3, 2) = 0
            lCosseguros(4, 1) = 0
            lCosseguros(4, 2) = 0
            lCosseguros(5, 1) = 0
            lCosseguros(5, 2) = 0

            lI = 0

            For Each lobjstcCoss In pobjstcPed.ColecaoPedCoss
                lI = lI + 1
                lCosseguros(lI, 1) = lobjstcCoss.Cod_Coss
                lCosseguros(lI, 2) = lobjstcCoss.Per_Coss
            Next

            .Pro_Cosseguros = lCosseguros
            .Pro_DebAgencia = pobjstcPed.gAgenciaDig
            .Pro_DebBanco = pobjstcPed.Cod_Bco_Cobr
            .Pro_DebConta = pobjstcPed.Num_Cta_Bas_Cor

            'Abrindo form para valida��o da ag�ncia e conta.
            If pobjstcPed.Num_Cta_Bas_Cor <> 0 Then
                Screen.MousePointer = vbDefault

                If gfCarCboBancos(3, frmT46V110A.cboDebBanco, lMensagem) = False Then
                    Call gM46V999.gpGraLog(2, lMensagem)
                    Exit Sub
                End If

                frmT46V110A.cboDebBanco.ListIndex = -1

                lI = 0

                For lI = 0 To (frmT46V110A.cboDebBanco.ListCount - 1)
                    If pobjstcPed.Cod_Bco_Cobr = frmT46V110A.cboDebBanco.ItemData(lI) Then
                        frmT46V110A.cboDebBanco.ListIndex = lI
                        Exit For
                    End If
                Next lI

                lDigito = 0

                Set clsA99V001 = New clsA99V001
                Set clsA99V001.M46V999 = gM46V999

                Select Case pobjstcPed.Cod_Bco_Cobr
                Case 237
                    Call clsA99V001.DIG_BRADESCO(Format(pobjstcPed.Num_Cta_Bas_Cor, "0000000"), lDigito, lAuxiliar)
                Case 341
                    Call clsA99V001.DIG_ITAU(Format(pobjstcPed.gAgenciaDig, "0000"), Format(pobjstcPed.Num_Cta_Bas_Cor, "00000"), lDigito, lAuxiliar)
                Case 347
                    If pobjstcPed.gAgenciaDig >= 1000 Then
                        Call clsA99V001.DIG_SUDAMERIS_NOVO(Format(pobjstcPed.gAgenciaDig, "0000"), Format(pobjstcPed.Num_Cta_Bas_Cor, "0000000"), lDigito, lAuxiliar)
                    Else
                        Call clsA99V001.DIG_SUDAMERIS(Format(pobjstcPed.gAgenciaDig, "000"), Format(pobjstcPed.Num_Cta_Bas_Cor, "000000000"), lDigito, lAuxiliar)
                    End If
                Case 409
                    Call clsA99V001.DIG_UNIBANCO(Format(pobjstcPed.gAgenciaDig, "0000"), Format(pobjstcPed.Num_Cta_Bas_Cor, "000000"), lDigito, lAuxiliar)
                Case 479
                    Call clsA99V001.DIG_BOSTON(Format(pobjstcPed.Num_Cta_Bas_Cor, "0000000"), lDigito, lAuxiliar)
                Case 356
                    Call clsA99V001.DIG_REAL(Format(pobjstcPed.gAgenciaDig, "0000"), Format(pobjstcPed.Num_Cta_Bas_Cor, "0000000"), lDigito, lAuxiliar)
                Case 399
                    Call clsA99V001.DIG_HSBC(Format(pobjstcPed.gAgenciaDig, "0000"), Format(pobjstcPed.Num_Cta_Bas_Cor, "0000000000"), lDigito, lAuxiliar)
                Case 47
                    Call clsA99V001.DIG_BANESE(Format(pobjstcPed.gAgenciaDig, "0000"), Format(pobjstcPed.Num_Cta_Bas_Cor, "0000000000"), lDigito, lAuxiliar)
                Case 33
                    Call clsA99V001.DIG_SANTANDER(Format(pobjstcPed.gAgenciaDig, "0000"), Format(pobjstcPed.Num_Cta_Bas_Cor, "0000000000"), lDigito, lAuxiliar)
                'Novo Conv�nio banco BRB
                Case 70
                    Call clsA99V001.DIG_BANCOBRB(Format(pobjstcPed.gAgenciaDig, "0000"), Format(pobjstcPed.Num_Cta_Bas_Cor, "000000000"), lDigito, lAuxiliar)
                End Select

                frmT46V110A.txtDebAgencia.Text = pobjstcPed.gAgenciaDig    '1001180 - Valida��o de DV na Triagem
                If pobjstcPed.Cod_Bco_Cobr = 399 Then
                    frmT46V110A.txtDebConta.Text = Left(Format(pobjstcPed.Num_Cta_Bas_Cor, "0000000000"), 9) & "-" & Right(Format(pobjstcPed.Num_Cta_Bas_Cor, "0000000000"), 1) & lDigito
                Else
                    frmT46V110A.txtDebConta.Text = pobjstcPed.Num_Cta_Bas_Cor & "-" & lDigito
                End If

                Set clsA99V001 = Nothing

                gProChamado = "T46V012A"

                Call frmT46V110A.mpInstanciarForm(2)

                Screen.MousePointer = vbHourglass

                'Salvo Banco, Ag�ncia e Conta'
                Set clsA99V001 = Nothing

                pobjstcPed.Cod_Bco_Cobr = Mid(frmT46V110A.cboDebBanco.Text, 1, 3)

                '1001180 - Valida��o de DV na Triagem
                'Corre��o para banco Bradesco
                If pobjstcPed.Cod_Bco_Cobr = 237 Then
                    pobjstcPed.Cod_Agen_Bas_Cor = Left(Format(val(Replace(frmT46V110A.txtDebAgencia.Text, "-", "")), "00000"), 4)
                    pobjstcPed.Cod_Dig_Agen_Bas_Cor = Right(Format(val(Replace(frmT46V110A.txtDebAgencia.Text, "-", "")), "00000"), 1)
                Else
                    pobjstcPed.Cod_Agen_Bas_Cor = val(Replace(frmT46V110A.txtDebAgencia.Text, "-", ""))
                    pobjstcPed.Cod_Dig_Agen_Bas_Cor = 0
                End If
                If pobjstcPed.Cod_Bco_Cobr = 399 Then
                    pobjstcPed.Num_Cta_Bas_Cor = Replace(frmT46V110A.txtDebConta.Text, "-", "")
                    pobjstcPed.Num_Cta_Bas_Cor = Left(Format(pobjstcPed.Num_Cta_Bas_Cor, "00000000000"), 10)

                Else
                    pobjstcPed.Num_Cta_Bas_Cor = Mid(Trim(frmT46V110A.txtDebConta.Text), 1, Len(Trim(frmT46V110A.txtDebConta.Text)) - 2)
                End If

                Set lObjPed = New clsA46V702A
                Set lObjPed.M46V999 = gM46V999

                pobjstcPed.Dat_Liber = Format(Date, "yyyymmdd")
                pobjstcPed.Cod_User_Liber = gM46V999.gNomUsuario6

                If Not lObjPed.mfAlterar(pobjstcPed, "", "", mUltDatAlt, mUltUsuario) Then
                    Call mpInformarErro(lObjPed.UltimoErro, 2, True)
                    Exit Sub
                End If
                mUltDatAlt = pobjstcPed.Dat_Liber
                mUltUsuario = pobjstcPed.Cod_User_Liber

                .Pro_DebConta = Replace(Trim(frmT46V110A.txtDebConta.Text), "-", "")

                Unload frmT46V110A

            End If

            'Restitui��o
            .Pro_CreBanco = pobjstcPed.objstcCtrlEmis.Cod_Banco_Cre
            .Pro_CreAgencia = pobjstcPed.objstcCtrlEmis.Cod_Agenc_Cre
            .Pro_CreConta = pobjstcPed.objstcCtrlEmis.Cod_Conta_Cre & pobjstcPed.objstcCtrlEmis.Cod_DIG_Conta
            .Pro_Tip_Res_Cre = pobjstcPed.objstcCtrlEmis.Tip_Res_Cre

            .Pro_End_Bairro = pobjstcPed.objstcPedEndereco.Nom_Bairro
            .Pro_End_CEP = Format(pobjstcPed.objstcPedEndereco.Num_CEP, "00000000")
            .Pro_End_Cidade = pobjstcPed.objstcPedEndereco.Nom_Cid

            If gM46V999.gfPreenchido(pobjstcPed.objstcPedEndereco.Dsc_Compl_Logr) Then
                .Pro_End_Complemento = pobjstcPed.objstcPedEndereco.Dsc_Compl_Logr
            Else
                .Pro_End_Complemento = ""
            End If

            .Pro_End_NomLogradouro = pobjstcPed.objstcPedEndereco.Nom_Logr
            .Pro_End_TipLogradouro = pobjstcPed.objstcPedEndereco.Cod_Tip_Logr
            .Pro_End_UF = pobjstcPed.objstcPedEndereco.Sig_UF

            'Proposta: forma de pagamento.
            '1 = carn� e 2 = d�bito em conta.
            If pobjstcPed.Num_Cta_Bas_Cor <> 0 And pobjstcPed.Num_Cta_Bas_Cor <> 9999999 Then
                .Pro_ForPagamento = 2
            ElseIf pobjstcPed.Num_Cta_Bas_Cor = 9999999 And pobjstcPed.Cod_Agen_Bas_Cor = 9999 And pobjstcPed.Cod_Bco_Cobr = 999 Then
                .Pro_ForPagamento = 4
            Else
                .Pro_ForPagamento = 1
            End If

            .Pro_IndAgencia = pobjstcPed.gAgenciaDig
            .Pro_IndBanco = pobjstcPed.Cod_Bco_Cobr
            .Pro_IndConta = pobjstcPed.Num_Cta_Bas_Cor

            .Pro_IndCobranca = 2
            .Pro_IndCosseguros = IIf(pobjstcPed.ColecaoPedCoss.Count > 0, 1, 2)
            .Pro_NomSegurado = pobjstcPed.Nom_Segurado
            .Pro_PagAntBloqueto = pobjstcPed.objstcCtrlEmis.Num_Bloq

            If pobjstcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec = e_SimNao_Sim Then
                If pobjstcPed.objstcCtrlEmis.Num_Bloq <> 0 Then
                    .Pro_PagAntecipado = 1
                Else
                    .Pro_PagAntecipado = 2
                End If
            Else
                .Pro_PagAntecipado = 0
            End If
            .Pro_QtdParcelas = pobjstcPed.Qtd_Parc
            .Pro_UniProdutiva = pobjstcPed.Cod_Unid_Prod
            .Pro_VigInicio = pobjstcPed.Dat_Ini_Vig
            .Pro_TipPessoa = pobjstcPed.Tip_Pessoa
            .Pro_CNPJ_CPF = pobjstcPed.Num_Cgc_Cpf
            .Pro_CodSegurado = pobjstcPed.Cod_Segurado

            'b) Executa consist�ncia.
            On Error Resume Next

            Call .VER_COMUM_NIVEL_2

            lI = 0

            For lI = 1 To gM46V999.gQuaErros
                Call mpInformarErro(CStr(gM46V999.gerros(lI)), 2, True)
                lI = lI + 1
            Next

            If Err <> 0 Then
                Call mpInformarErro("FTE0009 - " & gM46V999.gPrefixo & "Erro ao consistir dados comuns n�vel 2 (" & _
                                    Err & " - " & Error & ").", 2, True)
                On Error GoTo 0
                Exit Sub
            End If

            On Error GoTo 0

            Dim lZerar_campoCodigo As Boolean: lZerar_campoCodigo = False
            If Left(pobjstcPed.Cod_Segurado, 4) = 7777 And pobjstcPed.objstcCtrlEmis.Cod_Corr <> 1601 Then
                lZerar_campoCodigo = True
            Else
                Dim bdGed As ADODB.Connection
                Dim lrst As clsYasRecordSet
                Dim lSelect As String
                Dim lCpfCnpj As String

                '[Verifica se PF/PJ]
                lCpfCnpj = pobjstcPed.Num_Cgc_Cpf

                '[Abrir base GED.tab_cliente]
                If Not gM46V999.gfAbrBasDados("GED", gM46V999.gTipGed, gM46V999.gEndGed, bdGed, lMensagem) Then Call gM46V999.gpGraLog(1, lMensagem)

                '[Verificar se CPF/CNPJ dif. de 0 e C�d.= 0]
                If lCpfCnpj <> "0" And pobjstcPed.Cod_Segurado = 0 Then

                    '[Obter os C�digos cadastrados na GED de acordo com o CPF/CNPJ digitado]
                    lSelect = "SELECT COD_CLI FROM tab_cliente where CGC_CPF='" & Trim(lCpfCnpj) & "'"
                    If Not gM46V999.gfObtRegistro(bdGed, lSelect, lrst, lMensagem) Then Call gM46V999.gpGraLog(1, lMensagem)

                    '[Caso v�rios c�digos retornados, selecionar o primeiro]
                    If Not (lrst.EOF And lrst.BOF) Then
                        lrst.MoveLast: lrst.MoveFirst
                        If lrst.RecordCount = 1 Then
                            pobjstcPed.Cod_Segurado = val(lrst("Cod_Cli"))
                        End If
                    End If

                    '[Verificar se c�d. dif. de 0 no campo]
                ElseIf pobjstcPed.Cod_Segurado <> 0 Then

                    '[Obter os CPFs/CNPJs cadastrados na GED com o c�digo digitado]
                    lSelect = "SELECT * FROM tab_cliente where COD_CLI=" & pobjstcPed.Cod_Segurado
                    If Not gM46V999.gfObtRegistro(bdGed, lSelect, lrst, lMensagem) Then Call gM46V999.gpGraLog(1, lMensagem)

                    '[C�digo inv�lido:]
                    If lrst.EOF And lrst.BOF Then
                        lZerar_campoCodigo = True
                    End If
                End If
            End If

            If lZerar_campoCodigo Then
                pobjstcPed.Cod_Segurado = 0
            End If
            '[Limpar Mem�ria]
            lSelect = ""
            Set lrst = Nothing
            Set bdGed = Nothing
        End With

        Set clsA46V102 = Nothing
        Set lObjPed = New clsA46V702A
        Set lObjPed.M46V999 = gM46V999

        Dim mDatTerVig As Date
        Dim lAuxMsg As String
        Dim lobjPedObs As clsA46V714A     'Objeto da classe de observa��o.

        'Consiste vig�ncia e renova��o
        Select Case pobjstcPedLoc.Tip_Emissao
        Case e_TipEmissao_RenYasuda
            mDatTerVig = gM46V999.gfForData(2, lObjPed.mfRetornaTerminoVigencia(pobjstcPed))

            If Mid(mDatTerVig + 30, 7, 4) & Mid(mDatTerVig + 30, 4, 2) & Mid(mDatTerVig + 30, 1, 2) < pobjstcPed.Dat_Ini_Vig Then
                Call mpInformarErro("In�cio de vig�ncia superior a 30 dias " & _
                                    "do final de vig�ncia da ap�lice anterior! Favor calcular seguro novo.", 2, True)
            End If
        End Select

        '3. Consiste dados de identifica��o do proponente.
        If pobjstcPed.objstcCtrlEmis.Cod_Iden_Pessoa = e_TipPessoa_Fisica Then
            lTipPessoa = 2  'Pessoa f�sica.
        Else
            lTipPessoa = 1  'Pessoa jur�dica.
        End If

        If pobjstcPed.objstcCtrlEmis.Num_Bloq & pobjstcPed.objstcCtrlEmis.DIG_BLOQUETO <> 0 Then
            lAuxiliar = pobjstcPed.objstcCtrlEmis.Num_Bloq & pobjstcPed.objstcCtrlEmis.DIG_BLOQUETO
            Set clsA99V001 = New clsA99V001
            Set clsA99V001.M46V999 = gM46V999

            Call clsA99V001.DIG_BLOQUETO(Format(lAuxiliar, "000000000"), lDigito, lMensagem)
            If Right(lAuxiliar, 1) <> lDigito Then
                Call mpInformarErro("Pagamento antecipado com bloqueto - n�mero do bloqueto com d�gito inv�lido.", 2, True)
            End If
        End If

        Dim lLimDias As Integer
        lLimDias = 31

        'Ficha 1000837 - Triagem Informar Beneficiario nos Itens
        '        Dim lLinha As Integer
        '        Dim lTotLin As Long
        '        Dim lSeq As Long
        '        Dim lCod_Espec As Integer
        '        Dim AuxArray As Variant
        '        Dim lTextoInc As String
        '        Dim lobjstcTextoEspec As stcA46V726B
        '        Dim lobjTextoEspec As clsA46V726A
        '        Dim lNum_Pag As Integer
        '        Dim i As Integer
        '
        '        If MsgBox("Deseja informar benefici�rio(s) para o(s) item(ns)?", vbQuestion + vbYesNo) = vbYes Then
        '            Screen.MousePointer = vbNormal
        '            frmT46V166A.flxEspec.Rows = 1: frmT46V166A.flxEspec.Rows = 2
        '            lI = 1
        '            Set lobjTextoEspec = New clsA46V726A
        '            Set lobjTextoEspec.M46V999 = gM46V999
        '
        '            If pobjstcPed.ColecaoTextoEspec.Count > 0 Then
        '                For Each lobjstcTextoEspec In pobjstcPed.ColecaoTextoEspec
        '                    With frmT46V166A
        '                        .txtTextoEspec = ""
        '                        .txtTextoEspec = lobjstcTextoEspec.Espec_lin1 & vbCrLf
        '                        .txtTextoEspec = .txtTextoEspec & lobjstcTextoEspec.Espec_lin2 & vbCrLf
        '                        .txtTextoEspec = .txtTextoEspec & lobjstcTextoEspec.Espec_lin3 & vbCrLf
        '                        .txtTextoEspec = .txtTextoEspec & lobjstcTextoEspec.Espec_lin4 & vbCrLf
        '                        .txtTextoEspec = .txtTextoEspec & lobjstcTextoEspec.Espec_lin5 & vbCrLf
        '                        .txtTextoEspec = .txtTextoEspec & lobjstcTextoEspec.Espec_lin6 & vbCrLf
        '                        .txtTextoEspec = .txtTextoEspec & lobjstcTextoEspec.Espec_lin7 & vbCrLf
        '                        .txtTextoEspec = .txtTextoEspec & lobjstcTextoEspec.Espec_lin8 & vbCrLf
        '                        .txtTextoEspec = .txtTextoEspec & lobjstcTextoEspec.Espec_lin9 & vbCrLf
        '                        .txtTextoEspec = .txtTextoEspec & lobjstcTextoEspec.Espec_lin10 & vbCrLf
        '                        .txtTextoEspec = .txtTextoEspec & lobjstcTextoEspec.Espec_lin11 & vbCrLf
        '                        .txtTextoEspec = .txtTextoEspec & lobjstcTextoEspec.Espec_lin12 & vbCrLf
        '                        .txtTextoEspec = .txtTextoEspec & lobjstcTextoEspec.Espec_lin13 & vbCrLf
        '                        .txtTextoEspec = .txtTextoEspec & lobjstcTextoEspec.Espec_lin14 & vbCrLf
        '                        .txtTextoEspec = .txtTextoEspec & lobjstcTextoEspec.Espec_lin15 & vbCrLf
        '                        .txtTextoEspec = .txtTextoEspec & lobjstcTextoEspec.Espec_lin16 & vbCrLf
        '                        .txtTextoEspec = .txtTextoEspec & lobjstcTextoEspec.Espec_lin17 & vbCrLf
        '                        .txtTextoEspec = .txtTextoEspec & lobjstcTextoEspec.Espec_lin18 & vbCrLf
        '                        .txtTextoEspec = .txtTextoEspec & lobjstcTextoEspec.Espec_lin19 & vbCrLf
        '                        .txtTextoEspec = .txtTextoEspec & lobjstcTextoEspec.Espec_lin20
        '                        .flxEspec.Rows = lI + 1
        '                        .flxEspec.TextMatrix(lI, 0) = lI
        '                        .flxEspec.TextMatrix(lI, 1) = lobjstcTextoEspec.Cod_Espec
        '                        .flxEspec.TextMatrix(lI, 2) = .txtTextoEspec
        '                        lI = lI + 1
        '                        If Not lobjTextoEspec.mfExcluir(lobjstcTextoEspec, "", "") Then
        '                            lMensagem = lobjTextoEspec.UltimoErro
        '                            Call mpInformarErro(lMensagem, 2, True)
        '                        Else
        '                            pobjstcPed.ColecaoTextoEspec.Remove 1
        '                        End If
        '                    End With
        '                Next lobjstcTextoEspec
        '            End If
        '
        '            frmT46V166A.Show vbModal
        '            lNum_Pag = 1
        '            Screen.MousePointer = vbHourglass
        '            With frmT46V166A
        '                For lLinha = 1 To .flxEspec.Rows - 1
        '                    lSeq = 0
        '                    lTotLin = 1
        '                    If gM46V999.CDblx(.flxEspec.TextMatrix(lLinha, 1)) <> 0 Then
        '                        lCod_Espec = .flxEspec.TextMatrix(lLinha, 1)
        '                        If InStr(.flxEspec.TextMatrix(lLinha, 2), vbCrLf) <> 0 Then
        '                            AuxArray = Split(.flxEspec.TextMatrix(lLinha, 2), vbCrLf)
        '                            lTextoInc = AuxArray(lSeq)
        '                            lTotLin = UBound(AuxArray)
        '                        Else
        '                            lSeq = 1
        '                            lTextoInc = .flxEspec.TextMatrix(lLinha, 2)
        '                        End If
        '                        Set lobjstcTextoEspec = New stcA46V726B
        '                        With lobjstcTextoEspec
        '                            .NOSSO_NUMERO = pobjstcPed.NOSSO_NUMERO
        '                            .Cod_Org = 0
        '                            .Cod_Ramo = gM46V999.gCod_Ramo
        '                            .Num_Ped = pobjstcPed.Num_Ped
        '                            .Num_Apol = 0
        '                            .Num_Endo = 0
        '                            .Cod_Espec = lCod_Espec
        '                            .Num_pag = lNum_Pag
        '                            .Qtd_vias = 3
        '                            For i = 1 To lTotLin + 1
        '                                If i = 1 Then .Espec_lin1 = lTextoInc
        '                                If i = 2 Then .Espec_lin2 = lTextoInc
        '                                If i = 3 Then .Espec_lin3 = lTextoInc
        '                                If i = 4 Then .Espec_lin4 = lTextoInc
        '                                If i = 5 Then .Espec_lin5 = lTextoInc
        '                                If i = 6 Then .Espec_lin6 = lTextoInc
        '                                If i = 7 Then .Espec_lin7 = lTextoInc
        '                                If i = 8 Then .Espec_lin8 = lTextoInc
        '                                If i = 9 Then .Espec_lin9 = lTextoInc
        '                                If i = 10 Then .Espec_lin10 = lTextoInc
        '                                If i = 11 Then .Espec_lin11 = lTextoInc
        '                                If i = 12 Then .Espec_lin12 = lTextoInc
        '                                If i = 13 Then .Espec_lin13 = lTextoInc
        '                                If i = 14 Then .Espec_lin14 = lTextoInc
        '                                If i = 15 Then .Espec_lin15 = lTextoInc
        '                                If i = 16 Then .Espec_lin16 = lTextoInc
        '                                If i = 17 Then .Espec_lin17 = lTextoInc
        '                                If i = 18 Then .Espec_lin18 = lTextoInc
        '                                If i = 19 Then .Espec_lin19 = lTextoInc
        '                                If i = 20 Then .Espec_lin20 = lTextoInc
        '                                If lSeq < lTotLin And lSeq + 1 < 20 Then
        '                                    lSeq = lSeq + 1
        '                                    lTextoInc = AuxArray(lSeq)
        '                                Else
        '                                    Exit For
        '                                End If
        '                            Next
        '                            .Num_PI = gM46V111.gNum_Pi
        '                        End With
        '                        If Not lobjTextoEspec.mfIncluir(lobjstcTextoEspec, "", "") Then
        '                            lMensagem = lobjTextoEspec.UltimoErro
        '                            Call mpInformarErro(lMensagem, 2, True)
        '                        Else
        '                            pobjstcPed.ColecaoTextoEspec.Add lobjstcTextoEspec, lobjstcTextoEspec.Class_ID
        '                        End If
        '                        lNum_Pag = lNum_Pag + 1
        '                    End If
        'Proximo:
        '                Next
        '                .flxEspec.Rows = 1: .flxEspec.Rows = 2
        '            End With
        '        End If


        '                                       [Ajuste para Depto de Triagem]
        If pobjstcPed.ColecaoPedVenc.Count <> 0 Or Left(pobjstcPed.Cod_Segurado, 4) = 7777 Then
            Screen.MousePointer = vbDefault
            gProChamado = "T46V012A"
            Call frmT46V121A.gpInstanciarForm(pobjstcPed.Qtd_Parc, pobjstcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec <> e_SimNao_Nao, pobjstcPed.ColecaoPedVenc)

            Screen.MousePointer = vbHourglass
            'Salvar na cole��o

            'Salvo as datas de vencimento ajustadas
            Set objTabVenc = New clsA46V713A
            Set objTabVenc.M46V999 = gM46V999

            Set lObjstcVenc = New stcA46V713B
            Set lObjstcVenc.M46V999 = gM46V999

            lObjstcVenc.NOSSO_NUMERO = pobjstcPed.NOSSO_NUMERO
            If Not objTabVenc.mfExcluir(lObjstcVenc, "", "", "") Then
                Call mpInformarErro(objTabVenc.UltimoErro, 2, True)
            Else
                Set pobjstcPed.ColecaoPedVenc = New Collection
                For lI = 1 To (pobjstcPed.Qtd_Parc)
                    If IsDate(frmT46V121A.txtDtVenc(lI).Text) = False Then
                        lMensagem = "FGE0082 - Data de vencimento da parcela " & lI + 1 & " inv�lida (n�o preenchida)."
                        Call mpInformarErro(lMensagem, 2, True)
                    Else
                        Set lObjstcVenc = New stcA46V713B
                        Set lObjstcVenc.M46V999 = gM46V999

                        lObjstcVenc.NOSSO_NUMERO = pobjstcPed.NOSSO_NUMERO
                        lObjstcVenc.Cod_Ramo = pobjstcPed.Cod_Ramo
                        lObjstcVenc.Num_Ped = pobjstcPed.Num_Ped
                        lObjstcVenc.Num_Parc = lI
                        lObjstcVenc.DAT_VENC = gM46V999.gfForData(1, frmT46V121A.txtDtVenc(lI).Text)
                        If Not objTabVenc.mfIncluir(lObjstcVenc, "", "", "") Then
                            lMensagem = objTabVenc.UltimoErro
                            Call mpInformarErro(lMensagem, 2, True)
                        Else
                            pobjstcPed.ColecaoPedVenc.Add lObjstcVenc, lObjstcVenc.Class_ID
                        End If
                    End If
                Next
            End If
            Unload frmT46V121A
        Else
            If (UCase(pobjstcPedLoc.objstcPedObs.Observacao) Like "*VENC*" And Not UCase(pobjstcPedLoc.objstcPedObs.Observacao) Like "*VENCIDO*") Or _
               UCase(pobjstcPedLoc.objstcPedObs.Observacao) Like "*PARC*" Or _
               UCase(pobjstcPedLoc.objstcPedObs.Observacao) Like "*PAGAMENTO*" & _
               UCase(pobjstcPedLoc.objstcPedObs.Observacao) Like "*PAG*" Or _
               (UCase(pobjstcPedLoc.objstcPedObs.Observacao) Like "*DIA*" And Not UCase(pobjstcPedLoc.objstcPedObs.Observacao) Like "*VENCIDO*") Or _
               UCase(pobjstcPedLoc.objstcPedObs.Observacao) Like "*DEB*" Or _
               UCase(pobjstcPedLoc.objstcPedObs.Observacao) Like "*PAGTO*" Then
                Screen.MousePointer = vbDefault
                gProChamado = "T46V012"
                Call frmT46V121A.gpInstanciarForm(pobjstcPed.Qtd_Parc, pobjstcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec <> e_SimNao_Nao, pobjstcPed.ColecaoPedVenc)

                Screen.MousePointer = vbHourglass
                'Salvar na cole��o
                Set pobjstcPed.ColecaoPedVenc = New Collection
                If gM46V999.gfPreenchido(frmT46V121A.txtDtVenc(1).Text) = True Then
                    For lI = 1 To (pobjstcPed.Qtd_Parc)
                        If IsDate(frmT46V121A.txtDtVenc(lI).Text) = False Then
                            lMensagem = "FGE0082 - Data de vencimento da parcela " & lI & " inv�lida (n�o preenchida)."
                            Call mpInformarErro(lMensagem, 2, True)
                        ElseIf gM46V999.gfForData(1, frmT46V121A.txtDtVenc(lI).Text) > pobjstcPed.Dat_Ter_Vig Then
                            lMensagem = "FGE0082 - Data de vencimento da parcela " & lI & " inv�lida (maior que o t�rmino da vig�ncia)."
                            Call mpInformarErro(lMensagem, 2, True)
                        ElseIf gM46V999.gfForData(1, frmT46V121A.txtDtVenc(lI).Text) < gM46V999.gHoje Then
                            lMensagem = "FGE0082 - Data de vencimento da parcela " & lI & " inv�lida (menor que a data atual)."
                            Call mpInformarErro(lMensagem, 2, True)
                        Else
                            Set lObjstcVenc = New stcA46V713B
                            Set lObjstcVenc.M46V999 = gM46V999

                            lObjstcVenc.NOSSO_NUMERO = pobjstcPed.NOSSO_NUMERO
                            lObjstcVenc.Cod_Ramo = pobjstcPed.Cod_Ramo
                            lObjstcVenc.Num_Ped = pobjstcPed.Num_Ped
                            lObjstcVenc.Num_Parc = lI
                            lObjstcVenc.DAT_VENC = gM46V999.gfForData(1, frmT46V121A.txtDtVenc(lI).Text)
                            Set objTabVenc = New clsA46V713A
                            Set objTabVenc.M46V999 = gM46V999

                            If Not objTabVenc.mfIncluir(lObjstcVenc, "", "", "") Then
                                lMensagem = objTabVenc.UltimoErro
                                Call mpInformarErro(lMensagem, 2, True)
                            Else
                                pobjstcPed.ColecaoPedVenc.Add lObjstcVenc, lObjstcVenc.Class_ID
                            End If
                        End If
                    Next
                End If
                Unload frmT46V121A
            End If
        End If
    End If

    'If pobjstcPed.Cod_Bco_Cobr <> 341 And pobjstcPed.Num_Cta_Bas_Cor = 0 And pobjstcPed.Cod_Tip_Doc = e_TipDoc_Proposta Then
    '    If pobjstcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec = e_SimNao_Nao And pobjstcPed.Qtd_Parc = 1 Then
    '        'Ignorar regra do banco
    'antecipado com uma parcela
    '    Else
    '        pobjstcPed.Cod_Bco_Cobr = 341
    '    End If
    'End If

End Sub

Function mpVerificaUnicidade(pbdImport As ADODB.Connection, ByVal pNossoNumero As String, _
                             ByVal pDigVer As Long, ByRef pRetorno As Integer, _
                             ByRef pMensagem As String, Optional ByRef pCodUnic As Long) As Boolean
    Dim lSql As String
    Dim lrsTabela As clsYasRecordSet    'Registro da tabela.

    mpVerificaUnicidade = False

    lSql = "SELECT Cod_Unicidade FROM Ftp_Tab_Ped"
    lSql = lSql & " where Nosso_Numero = " & gM46V999.gfFormatarCamposSQL(pNossoNumero)

    If gM46V999.gfObtRegistro(pbdImport, lSql, lrsTabela, pMensagem) = False Then
        Call gM46V999.gpFecha2(pbdImport)
        Exit Function
    End If

    If lrsTabela.EOF Then
        pRetorno = 1
    Else
        Do While lrsTabela.EOF = False
            pCodUnic = lrsTabela("Cod_Unicidade")
            If gM46V999.Check(lrsTabela("Cod_Unicidade")) <> pDigVer Then
                pRetorno = 2
            Else
                'Encontrado transmiss�o com mesmo DV
                pRetorno = 3
                Exit Do
            End If
            lrsTabela.MoveNext
        Loop
    End If

    mpVerificaUnicidade = True

End Function

Private Sub mpVerIteNiv_2(pobjstcPed As stcA46V702B, pobjstcItem As stcA46V708B)
'Procedure: consiste item n�vel 2: consiste todos dados do item.

    Dim lobjA46V103A As clsA46V103A
    Dim objTabUser As clsA46V725A
    Dim objstcTabUser As stcA46V725B

    Set lobjA46V103A = New clsA46V103A
    Set lobjA46V103A.M46V103 = gM46V103
    
    Set objTabUser = New clsA46V725A
    Set objTabUser.M46V999 = gM46V999

    Set objstcTabUser = New stcA46V725B
    objstcTabUser.login_user_rede = UCase(gM46V999.gNomUsuario)
    If Not objTabUser.mfCarregarColecao(objstcTabUser) Then
        Call gM46V999.gpGraLog(1, objTabUser.UltimoErro)
        Exit Sub
    End If
    
    Call lobjA46V103A.gpVerIteNiv_2(pobjstcItem, pobjstcPed, val(pobjstcPed.Val_IS), True, objstcTabUser, "")

End Sub

Private Sub mpVerNormas(pobjstcPed As stcA46V702B, pobjstcPedLoc As stcA46V708B, pItem As Long, pCod_Retorno As Integer, pCod_Aceit As Integer, ByRef pMensInsp As String)
    Dim lretorno As String
    Dim lI As Integer
    Dim lCod As Integer
    Dim lAuxiliar As String
    Dim lBDP00Multi As ADODB.Connection
    Dim lRstErro As clsYasRecordSet
    Dim lMensagem As String
    Dim QtdErr As Integer
    Dim lSelect As String

    pMensInsp = ""
    QtdErr = 0
    If Not gM46V999.VerificarNormas Then
        Exit Sub
    End If
    lretorno = gM46V111.gfExecDLL(pobjstcPed, pobjstcPedLoc, e_TipProc_Normas, pItem, pCod_Retorno, gM46V999.gTipoHomologacao)

    If InStr(1, lretorno, "Existem erros de consist�ncia!") <> 0 Then
        'Busca os erros de consist�ncia.
        Call mpCapturaErros("Normas ")
    ElseIf gM46V999.gfPreenchido(lretorno) = True Then    'retornou com erros
        lI = InStr(1, lretorno, "--") + 2
        If gM46V999.gfAbrBasDados("P00MULTI2", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lBDP00Multi, lMensagem) = False Then
            Call gM46V999.gpGraLog(2, lMensagem)
            Exit Sub
        End If
        Do
            lCod = val(Mid(lretorno, lI, 3))
            'Se na tabela de parametriza��o retornou classe 1 despreza retornos 10 e 80
            If pCod_Aceit = 1 Then
                If lCod <> 55 Then
                    lI = InStr(lI, lretorno, "-") + 1
                    GoTo LerProximoErro
                End If
            End If
            lI = InStr(lI, lretorno, "-") + 1
            lAuxiliar = Mid(lretorno, lI, 11)
            lSelect = " Select * from Tab_Dsc_Retorno  where Codigo = " & val(lAuxiliar)

            If gM46V999.gfObtRegistro(lBDP00Multi, lSelect, lRstErro, lMensagem) = False Then
                Call gM46V999.gpGraLog(2, lMensagem)
                Exit Sub
            End If

            If Not lRstErro.EOF Then
                ' CR3923 - Item 1 - In�cio
                Dim mensagemInterna As String
                mensagemInterna = ""
                If Not lRstErro("descricao_interna") Is Nothing Then
                    mensagemInterna = Trim(lRstErro("descricao_interna"))
                End If

                If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda And mensagemInterna <> "" Then
                    lAuxiliar = val(lAuxiliar) & " - " & Trim(lRstErro("descricao_interna"))
                Else
                    lAuxiliar = val(lAuxiliar) & " - " & Trim(lRstErro("Descricao"))
                End If
                ' CR3923 - Item 1 - Fim
            End If

            If lCod = 55 Then
                pMensInsp = Trim(Mid(lAuxiliar, InStr(1, lAuxiliar, "-") + 1))
            Else
                Call mpInformarErro("Normas " & lAuxiliar, 2, True)
                QtdErr = QtdErr + 1
            End If

LerProximoErro:
            lI = InStr(lI, lretorno, "-") + 1
            If lI = 1 Then Exit Do
        Loop
        Call gM46V999.gpFecha2(lBDP00Multi)
    End If
    If QtdErr = 0 Then
        pCod_Retorno = 0
    End If
    DoEvents

End Sub

Private Sub mnuSair_Click()
    Unload Me
    If gM46V999.gfDisponivel("frmT46V001A") = True Then
        frmT46V001A.SetFocus
    End If
End Sub

Private Sub txtDigito_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub txtDV_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'Ficha 1000449 - Incluir campo Referencia Marsh para digita��o
        If gM46V999.gfPreenchido(txtNumeroPI.Text) And cboOpcao.ListIndex <> 2 And cboOpcao.ListIndex <> 3 Then
            Call mnuArqPesquisar_Click
        Else
            SendKeys vbTab
        End If
    End If
End Sub

Private Sub txtMTVVD_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If gM46V999.gfPreenchido(txtNumeroPI.Text) Then
            Call mnuArqPesquisar_Click
        Else
            SendKeys vbTab
        End If
    End If
End Sub

Private Sub txtNumeroPI_Change()
    txtPropostaSyas.Enabled = True
    txtDV.Enabled = True
End Sub

Private Sub txtNumeroPI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub txtNumeroPI_LostFocus()
    Dim lMensagem As String
    Dim lDigito As String
    Dim lobjCtrlEmis As clsA46V712A   'Objeto da classe clsA46V712A (TAB_CTRL_EMIS)
    Dim lobjstcCtrlEmis As stcA46V712B      'Objeto da classe de estrutura stcA46V712B (TAB_CTRL_EMIS)
    Dim bdBasDados As ADODB.Connection    'Base de dados.
    Dim lObjPed As clsA46V702A
    Dim lObjstcPed As stcA46V702B
    
    If TemOutroFormAberto = True Then Exit Sub

    If gM46V999.gfAbrBasDados("P00MPED", gM46V999.gTipP00MPED, gM46V999.gEndP00MPED, bdBasDados, _
                              lMensagem) = False Then
        Call gM46V999.gpGraLog(1, "FNP0044 - " & lMensagem)
        Exit Sub
    End If

    If Not IsNumeric(txtNumeroPI) Then Exit Sub
    If val(mAntPI) = val(txtNumeroPI.Text) Then Exit Sub

    mAntPI = txtNumeroPI.Text
    If Len(Trim(txtNumeroPI.Text)) <= 7 Then
        txtNumeroPI.Text = Right(Year(Date), 2) & "2" & Format(txtNumeroPI.Text, "0000000")
    Else
        txtNumeroPI.Text = Format(txtNumeroPI.Text, "0000000000")
    End If

    'Tenta localizar os dados no tab_ctrl_emis pelo PI.
    'Instancio os objetos
    Set lobjCtrlEmis = New clsA46V712A
    Set lobjCtrlEmis.M46V999 = gM46V999

    Set lobjstcCtrlEmis = New stcA46V712B

    Set lobjCtrlEmis.ConexaoBD = bdBasDados

    lobjstcCtrlEmis.NUM_PI = txtNumeroPI

    If Not lobjCtrlEmis.mfCarregarColecao(lobjstcCtrlEmis, "", "") Then
        Call gM46V999.gpGraLog(1, "FNP0043 - " & lobjCtrlEmis.UltimoErro)
        Exit Sub
    End If
    If lobjCtrlEmis.Colecao.Count <> 0 Then
        Set lobjstcCtrlEmis = lobjCtrlEmis.Colecao.Item(1)
        txtPropostaSyas = lobjstcCtrlEmis.NOSSO_NUMERO
        'If Not IsDate(txtDataRecebimento) Then
        If lobjstcCtrlEmis.Dat_Rec_Cia <> 0 Then
            txtDataRecebimento = gM46V999.gfForData(2, lobjstcCtrlEmis.Dat_Rec_Cia)
        End If
        'Calcula d�gito.
        Set clsA99V001 = New clsA99V001
        Set clsA99V001.M46V999 = gM46V999

        If clsA99V001.DIG_YASUDA(txtNumeroPI.Text, lDigito, lMensagem) = False Then
            Exit Sub
        End If
        Set clsA99V001 = Nothing
        txtDigito.Text = Format(lDigito, "0")
        Set lObjPed = New clsA46V702A
        Set lObjPed.M46V999 = gM46V999

        Set lObjstcPed = New stcA46V702B
        Set lObjstcPed.M46V999 = gM46V999

        lObjstcPed.NOSSO_NUMERO = lobjstcCtrlEmis.NOSSO_NUMERO
        If Not lObjPed.mfCarregarColecao(lObjstcPed, "", "") Then
            Call gM46V999.gpGraLog(1, lObjPed.UltimoErro)
            Exit Sub
        End If
        If lObjPed.Colecao.Count <> 0 Then
            Set lObjstcPed = lObjPed.Colecao.Item(1)
            txtDV.Text = lObjstcPed.Cod_Unicidade
        Else
            txtDV.Text = "99999"
        End If

        txtPropostaSyas.Enabled = False
        txtDV.Enabled = False
        If txtPropostaSyas.Enabled = True Then
            txtDV.SetFocus
        Else
            If txtDataRecebimento.Visible = True Then
                txtDataRecebimento.SetFocus
            End If
        End If

        If Left(lobjstcCtrlEmis.Cod_Corr, 2) = 80 Then
            cboOpcao.ListIndex = 2
            txtMTVVD.Text = Right(lobjstcCtrlEmis.Num_Ref_Corr, 6)
            txtMTVVD.SetFocus
        End If
        If lobjstcCtrlEmis.Cod_Ramo = 710 Then
            cboOpcao.ListIndex = 4
        End If
    Else
        txtPropostaSyas.Enabled = True
        txtDV.Enabled = True
        Call mpLimparCampos
    End If

    Set lobjCtrlEmis = Nothing
    Set lobjstcCtrlEmis = Nothing
    Call gM46V999.gpFecha2(bdBasDados)
End Sub

Private Function mfBuscarDadosSiscad(pNum_Apol As Double, _
                                     ByRef pIncluirSiscad As Boolean, ByRef pAlterarSiscad As Boolean, _
                                     ByRef pVerSiscad As Boolean, pTipEmissao As e_TipEmissao, pMensagem As String) As Boolean
    Dim lCPF_CNPJ As String
    Dim bdRamosDiversos As ADODB.Connection
    Dim rsRamosDiversos As clsYasRecordSet
    Dim rsSyasCob As clsYasRecordSet
    Dim lNumPI As Long
    Dim lCepSegDif As Boolean
    Dim lCepCobDif As Boolean
    Dim lAuxiliar As String
    Dim lAuxiliar2 As String
    Dim lSiscadDiv As Boolean
    Dim lSelect As String

    On Error GoTo ControleErro

    pIncluirSiscad = False
    pAlterarSiscad = False

    Call mpLimparErros
    Call mpInformarErro("", 1, False)

    ''**************************OBSERVA��O********************
    ''******Endere�o do segurado est�o no tab_ped_ender*******
    ''******Endere�o de cobran�a est� na tab_ped**************
    ''********************************************************

    lCPF_CNPJ = mobjstcPed.Num_Cgc_Cpf

    If gM46V999.gfAbrBasDados("RAMOSDIVERSOS", gM46V999.gTipDIVERSOS, gM46V999.gEndDIVERSOS, bdRamosDiversos, pMensagem) = False Then
        Exit Function
    End If

    lSelect = " select * from p0044800.dbo.t0044800 Where "
    lSelect = lSelect & " num_cgc_cpf = '" & Trim(lCPF_CNPJ) & "'"

    If gM46V999.gfObtRegistro(bdRamosDiversos, lSelect, rsRamosDiversos, pMensagem) = False Then
        Call gM46V999.gpFecha2(bdRamosDiversos)
        Exit Function
    End If

    If Not rsRamosDiversos.EOF Then
        lSiscadDiv = False
        lAuxiliar = Replace(Replace(mobjstcPed.Nom_Segurado, "  ", " "), "  ", " ")

        'Ficha 901047
        '        lAuxiliar2 = Replace(Trim(Left(Trim(rsRamosDiversos("NOME_RAZAO")) & " " & Trim(rsRamosDiversos("NOM_LOCTRAB")), 60)), "  ", " ")
        lAuxiliar2 = Replace(Replace(rsRamosDiversos("NOME_RAZAO"), "  ", " "), "  ", " ")

        'Ficha 901363
        If Trim(lAuxiliar) <> Trim(lAuxiliar2) Then
            Call mpInformarErro("Nome do Segurado divergente - " & "'" & lAuxiliar & "' <> '" & lAuxiliar2 & "'", 2, True)
            lSiscadDiv = True
        End If

        lCepSegDif = False

        If Trim(mobjstcPed.Tip_Pessoa) <> Trim(rsRamosDiversos("TIP_PESSOA")) Then
            Call mpInformarErro("Tipo de pessoa divergente - " & _
                                "'" & mobjstcPed.Tip_Pessoa & "' <> '" & rsRamosDiversos("TIP_PESSOA") & "'", 2, True)
        End If

        If Format(mobjstcPed.objstcPedEndereco.Num_CEP, "00000000") <> Format(rsRamosDiversos("Num_CEP"), "00000000") Then
            Call mpInformarErro("CEP divergente - " & _
                                "'" & mobjstcPed.objstcPedEndereco.Num_CEP & "' <> '" & Trim(rsRamosDiversos("Num_CEP")) & "'", 2, True)
            lCepSegDif = True
        Else
            If Right(Format(mobjstcPed.objstcPedEndereco.Num_CEP, "00000000"), 3) = "000" Then
                lCepSegDif = True
            End If
        End If

        If Trim(mobjstcPed.objstcPedEndereco.Cod_Tip_Logr) <> Trim(Left(Trim(rsRamosDiversos("COD_TIP_LOGR")), 3)) Then
            If lCepSegDif = False Then
                mobjstcPed.objstcPedEndereco.Cod_Tip_Logr = Trim(Left(Trim(rsRamosDiversos("COD_TIP_LOGR")), 3))
            Else
                Call mpInformarErro("Tipo de logradouro divergente - " & _
                                    "'" & mobjstcPed.objstcPedEndereco.Cod_Tip_Logr & "' <> '" & Trim(Left(Trim(rsRamosDiversos("COD_TIP_LOGR")), 3)) & "'", 2, True)
            End If
        End If

        If Trim(mobjstcPed.objstcPedEndereco.Nom_Logr) <> Trim(rsRamosDiversos("NOM_LOGR")) Then
            If lCepSegDif = False Then
                mobjstcPed.objstcPedEndereco.Nom_Logr = Trim(rsRamosDiversos("NOM_LOGR"))
            Else
                Call mpInformarErro("logradouro divergente - " & _
                                    "'" & mobjstcPed.objstcPedEndereco.Nom_Logr & "' <> '" & Trim(rsRamosDiversos("NOM_LOGR")) & "'", 2, True)
            End If
        End If

        lAuxiliar = Replace(mobjstcPed.objstcPedEndereco.Dsc_Compl_Logr, "APTO", "AP")
        lAuxiliar2 = Replace(Trim(rsRamosDiversos("DSC_COMPL_LOGR")), "APTO", "AP")
        lAuxiliar = Replace(lAuxiliar, ".", "")
        lAuxiliar2 = Replace(lAuxiliar2, ".", "")
        lAuxiliar = Replace(lAuxiliar, "-", " ")
        lAuxiliar2 = Replace(lAuxiliar2, "-", " ")
        lAuxiliar = Replace(lAuxiliar, "/", " ")
        lAuxiliar2 = Replace(lAuxiliar2, "/", " ")
        lAuxiliar = Replace(Replace(lAuxiliar, "  ", " "), "  ", " ")
        lAuxiliar2 = Replace(Replace(lAuxiliar2, "  ", " "), "  ", " ")

        If lAuxiliar <> lAuxiliar2 Then
            Call mpInformarErro("Complemento divergente - " & _
                                "'" & mobjstcPed.objstcPedEndereco.Dsc_Compl_Logr & "' <> '" & Trim(rsRamosDiversos("DSC_COMPL_LOGR")) & "'", 2, True)
        End If

        If Trim(mobjstcPed.objstcPedEndereco.Nom_Bairro) <> Trim(rsRamosDiversos("NOM_BAIRRO")) Then
            If lCepSegDif = False Then
                mobjstcPed.objstcPedEndereco.Nom_Bairro = Trim(rsRamosDiversos("NOM_BAIRRO"))
            Else
                Call mpInformarErro("Bairro divergente - " & _
                                    "'" & mobjstcPed.objstcPedEndereco.Nom_Bairro & "' <> '" & Trim(rsRamosDiversos("NOM_BAIRRO")) & "'", 2, True)
            End If
        End If

        If Trim(mobjstcPed.objstcPedEndereco.Nom_Cid) <> Trim(rsRamosDiversos("NOM_CID")) Then
            If lCepSegDif = False Then
                mobjstcPed.objstcPedEndereco.Nom_Cid = Trim(rsRamosDiversos("NOM_CID"))
            Else
                Call mpInformarErro("Cidade divergente - " & _
                                    "'" & mobjstcPed.objstcPedEndereco.Nom_Cid & "' <> '" & Trim(rsRamosDiversos("NOM_CID")) & "'", 2, True)
            End If
        End If

        If Trim(mobjstcPed.objstcPedEndereco.Sig_UF) <> Trim(rsRamosDiversos("SIG_UF")) Then
            Call mpInformarErro("UF divergente - " & _
                                "'" & mobjstcPed.objstcPedEndereco.Sig_UF & "' <> '" & Trim(rsRamosDiversos("SIG_UF")) & "'", 2, True)
        End If

        lAuxiliar = mfRetornarNumero(mobjstcPed.objstcPedEndereco.Num_Fone)
        lAuxiliar2 = mfRetornarNumero(Trim(rsRamosDiversos("NUM_FONE")))

        If lAuxiliar = "" Then
            lAuxiliar = lAuxiliar2
            mobjstcPed.objstcPedEndereco.Num_Fone = lAuxiliar
        End If

        If lAuxiliar <> Right(lAuxiliar2, Len(lAuxiliar)) Then
            If lAuxiliar Like "*" & Right(lAuxiliar2, Len(lAuxiliar)) & "*" Then
                mobjstcPed.objstcPedEndereco.Num_Fone = Trim(rsRamosDiversos("NUM_FONE"))
            Else
                If lAuxiliar2 Like "*11111*" Then
                    lSiscadDiv = True
                ElseIf lAuxiliar Like "*11111*" Then
                    mobjstcPed.objstcPedEndereco.Num_Fone = Trim(rsRamosDiversos("NUM_FONE"))
                Else
                    If gM46V999.gfPreenchido(lAuxiliar) Then
                        If Len(lAuxiliar) = 7 Then
                            lAuxiliar = Trim(rsRamosDiversos("NUM_DDD_FONE")) & lAuxiliar
                        End If
                        mobjstcPed.objstcPedEndereco.Num_Fone = lAuxiliar
                        lSiscadDiv = True
                    Else
                        Call mpInformarErro("Telefone divergente - " & _
                                            "'" & mobjstcPed.objstcPedEndereco.Num_Fone & "' <> '" & Trim(rsRamosDiversos("NUM_FONE")) & "'", 2, True)
                    End If
                End If
            End If
        End If

        If Trim(mobjstcPed.gLetraRamoAtiv) <> Trim(rsRamosDiversos("RAMO_ATIV_EMPR")) Then
            lSiscadDiv = True
        End If

        If IIf(mobjstcPed.Tip_Pessoa = 1, "J", "F") <> IIf(rsRamosDiversos("TIP_PESSOA") = 1, "J", "F") Then
            lSiscadDiv = True
        End If

        '[FRH] - 14/05/2009 - DE COD_TIP_DOC P/ COD_DOC
        If val(Trim(mobjstcPed.Cod_Doc)) <> val(Trim(rsRamosDiversos("TIP_DOCUM"))) Then
            lSiscadDiv = True
        End If

        If Trim(mobjstcPed.Num_Doc) <> Trim(rsRamosDiversos("NUM_DOCUM")) Then
            lSiscadDiv = True
        End If

        If Trim(mobjstcPed.Org_Exp) <> Trim(rsRamosDiversos("NOM_ORG_EXPED")) Then
            lSiscadDiv = True
        End If

        If mobjstcPed.Dat_Emis_Doc <> rsRamosDiversos("DAT_EXPED") Then
            lSiscadDiv = True
        End If

        'Ficha 900417
        If gM46V999.gCodProduto >= 968 Then
            If Trim(mobjstcPed.Nom_Empresa) <> Trim(rsRamosDiversos("NOM_PROFISSAO")) Then
                lSiscadDiv = True
            End If
        Else
            If Trim(mobjstcPed.Nom_Empresa) <> Trim(rsRamosDiversos("NOM_LOCTRAB")) Then
                lSiscadDiv = True
            End If
        End If

        If Trim(mobjstcPed.Sexo) <> Trim(rsRamosDiversos("SEXO")) Then
            lSiscadDiv = True
        End If

        If mobjstcPed.Val_RendaFatu <> gM46V999.gfTratarNulo(rsRamosDiversos("RendaFatu"), TpoValor) Then
            lSiscadDiv = True
        End If

        lNumPI = 0

        If pNum_Apol <> 0 Then
            'Pesquisar na tabela tab_ctrl_emis em qual PI a ap�lice q est� sendo renovada foi emitida...
            lSelect = " select top 1 * from ged.dbo.Tab_ctrl_emis Where "
            lSelect = lSelect & " num_apol = " & pNum_Apol
            lSelect = lSelect & " and cod_iden_stat_acao = 90"
            lSelect = lSelect & " order by dat_emis desc"

            If gM46V999.gfObtRegistro(bdRamosDiversos, lSelect, rsSyasCob, pMensagem) = False Then
                Call gM46V999.gpFecha2(bdRamosDiversos)
                Exit Function
            End If

            If Not rsSyasCob.EOF Then
                lNumPI = rsSyasCob("Num_PI")
            End If
        End If

        lSelect = " Select top 1 * "
        lSelect = lSelect & " from p0044800.dbo.T0044807 Where "
        lSelect = lSelect & " num_cgc_cpf = '" & Trim(lCPF_CNPJ) & "'"

        If lNumPI <> 0 Then
            lSelect = lSelect & " and num_pi =  " & lNumPI
        End If

        lSelect = lSelect & " order by dat_emis desc"

        If gM46V999.gfObtRegistro(bdRamosDiversos, lSelect, rsSyasCob, pMensagem) = False Then
            Call gM46V999.gpFecha2(bdRamosDiversos)
            Exit Function
        End If

        If Not rsSyasCob.EOF Then
            lCepCobDif = False

            If Format(mobjstcPed.Num_CEP, "00000000") <> Format(rsSyasCob("Num_CEP"), "00000000") Then
                Call mpInformarErro("CEP de cobran�a divergente - " & _
                                    "'" & mobjstcPed.Num_CEP & "' <> '" & rsSyasCob("Num_CEP") & "'", 2, True)
                lCepCobDif = True
            End If

            If Left(rsSyasCob("NOM_LOGR"), 4) = "FIN-" Then
                lCepCobDif = True
            End If

            If Trim(mobjstcPed.Nom_Logr) <> Trim(rsSyasCob("NOM_LOGR")) Then
                If lCepCobDif = False Then
                    mobjstcPed.Nom_Logr = Trim(rsSyasCob("NOM_LOGR"))
                Else
                    Call mpInformarErro("Logradouro de cobran�a divergente - " & _
                                        "'" & mobjstcPed.Nom_Logr & "' <> '" & Trim(rsSyasCob("NOM_LOGR")) & "'", 2, True)
                End If
            End If

            If Trim(mobjstcPed.Dsc_Compl_Logr) <> Trim(rsSyasCob("DSC_COMPL_LOGR")) Then
                If lCepCobDif = False Then
                    If mobjstcPed.Dsc_Compl_Logr Like "*" & Trim(rsSyasCob("DSC_COMPL_LOGR")) & "*" Or _
                       Trim(rsSyasCob("DSC_COMPL_LOGR")) Like "*" & mobjstcPed.Dsc_Compl_Logr & "*" Then
                        ''Altero o siscad com a informa��o da transmiss�o.
                    Else
                        Call mpInformarErro("Complemento de cobran�a divergente - " & _
                                            "'" & mobjstcPed.Dsc_Compl_Logr & "' <> '" & Trim(rsSyasCob("DSC_COMPL_LOGR")) & "'", 2, True)
                    End If
                Else
                    Call mpInformarErro("Complemento de cobran�a divergente - " & _
                                        "'" & mobjstcPed.Dsc_Compl_Logr & "' <> '" & Trim(rsSyasCob("DSC_COMPL_LOGR")) & "'", 2, True)
                End If
            End If

            If Trim(mobjstcPed.Nom_Bairro) <> Trim(rsSyasCob("NOM_BAIRRO")) Then
                If lCepCobDif = False Then
                    mobjstcPed.Nom_Bairro = Trim(rsSyasCob("NOM_BAIRRO"))
                Else
                    Call mpInformarErro("Bairro de cobran�a divergente - " & _
                                        "'" & mobjstcPed.Nom_Bairro & "' <> '" & Trim(rsSyasCob("NOM_BAIRRO")) & "'", 2, True)
                End If
            End If

            If Trim(mobjstcPed.Nom_Cid) <> Trim(rsSyasCob("NOM_CID")) Then
                If lCepCobDif = False Then
                    mobjstcPed.Nom_Cid = Trim(rsSyasCob("NOM_CID"))
                Else
                    Call mpInformarErro("Cidade de cobran�a divergente - " & _
                                        "'" & mobjstcPed.Nom_Cid & "' <> '" & Trim(rsSyasCob("NOM_CID")) & "'", 2, True)
                End If
            End If

            If Trim(mobjstcPed.Sig_UF) <> Trim(rsSyasCob("SIG_UF")) Then
                Call mpInformarErro("UF de cobran�a divergente - " & _
                                    "'" & mobjstcPed.Sig_UF & "' <> '" & rsSyasCob("SIG_UF") & "'", 2, True)
            End If
        End If

        If lSiscadDiv = True Or lCepSegDif = True Then
            pAlterarSiscad = True
        End If

        If pTipEmissao = e_TipEmissao_SeguroNovo Or pTipEmissao = e_TipEmissao_RenCongenere Then
            Call mpInformarErro("Validar siscad.", 2, True)
        End If
    Else
        pIncluirSiscad = True
        'Incluir os dados no siscad
        Call mpInformarErro("Validar siscad.", 2, True)
    End If


    'Ficha 1000831
    Dim lMsgNomDiferente As String
    Dim lListaNomes As String
    Dim RstAux As clsYasRecordSet

    lSelect = "Select * from ged.dbo.tab_ctrl_emis Where num_cgc_cpf = '" & Trim(mobjstcPed.Num_Cgc_Cpf) & "' and Num_PI <> " & val(txtNumeroPI.Text) & " order by dat_ini_vig desc"

    If Not gM46V999.gfObtRegistro(bdRamosDiversos, lSelect, RstAux, pMensagem) Then
        Call gM46V999.gpGraLog(1, pMensagem)
        Exit Function
    End If
    lMsgNomDiferente = ""
    Do While Not RstAux.EOF
        If Trim(UCase(RstAux("Nom_Seg"))) <> Trim(UCase(mobjstcPed.Nom_Segurado)) Then
            If Not lListaNomes Like "*;" & Trim(UCase(RstAux("Nom_Seg"))) & ";*" Then
                lMsgNomDiferente = lMsgNomDiferente & "PI: " & Format(RstAux("Num_PI"), "0000000000") & " - " & Trim(RstAux("Nom_Seg")) & vbCrLf
                lListaNomes = lListaNomes & ";" & Trim(UCase(RstAux("Nom_Seg"))) & ";"
            End If
        End If
        RstAux.MoveNext
    Loop
    If lMsgNomDiferente <> "" Then
        Screen.MousePointer = vbDefault
        MsgBox "Compara��o de titularidades de proponente/segurado divergente. Favor confirmar titularidade no site da Receita Federal." & vbCrLf & lMsgNomDiferente
        Screen.MousePointer = vbHourglass

        ''Fazer update
        lSelect = " INSERT INTO [GED].[dbo].[Tab_Comp_Titularidade]([Num_PI], [Nom_Proponente], [Dat_Comp], [Hor_Comp], [Msg_Comp], [Usuario])"
        lSelect = lSelect & " VALUES(" & val(txtNumeroPI.Text) & ",'" & mobjstcPed.Nom_Segurado & "'," & Format(Date, "yyyymmdd") & "," & Format(Time, "hhmmss") & ",'" & Left(Replace(lMsgNomDiferente, "'", "�"), 5000) & "','" & gM46V999.gNomUsuario & "')"

        If gM46V999.gfExeSQL(bdRamosDiversos, lSelect, pMensagem) <> 0 Then
            Call gM46V999.gpGraLog(1, pMensagem)
            Exit Function
        End If

    End If


    pVerSiscad = False

    If mQtdErros > 0 Then pVerSiscad = True
    mfBuscarDadosSiscad = True

    Exit Function

ControleErro:
    pMensagem = "Siscad - " & Err.Description

End Function

Private Sub mpConsisteItem(pobjstcPed As stcA46V702B, pobjstcPedLoc As stcA46V708B, pItem As Long, ByRef pCod_Retorno As Integer)
    Dim lretorno As String
    'Simony
    'N�o consiste Itens dos ramos diferente de Multi
    If gM46V111.gTipoGrupMulti <> e_GrupoRamo_Multi Then Exit Sub

    lretorno = gM46V111.gfExecDLL(pobjstcPed, pobjstcPedLoc, e_TipProc_ConsItem, pItem, pCod_Retorno, gM46V999.gTipoHomologacao)

    If InStr(1, lretorno, "Existem erros de consist�ncia!") <> 0 Then
        'Busca os erros de consist�ncia.
        Call mpCapturaErros("Consist�ncia do item " & pItem)
        pCod_Retorno = 0
    ElseIf gM46V999.gfPreenchido(lretorno) = True Then    'retornou sem erros
        Call mpInformarErro("Item " & Format(pItem, "00000") & "  " & lretorno, 2, True)
    End If

    DoEvents
End Sub


Private Function mfConsistirDados() As Boolean

    Dim lMensagem As String    'Mensagem
    Dim lDigito As String    'D�gito.
    Dim lDigitoNP As Integer    'D�gito.
    Dim lOpcaoSelecionada As Integer
    Dim bdBasDados As ADODB.Connection    'Base de dados.

    If gM46V999.gfAbrBasDados("P00MPED", gM46V999.gTipP00MPED, gM46V999.gEndP00MPED, bdBasDados, _
                              lMensagem) = False Then
        Call gM46V999.gpGraLog(1, lMensagem)
        Exit Function
    End If

    lOpcaoSelecionada = cboOpcao.ItemData(cboOpcao.ListIndex)

    txtDigito.Text = Format(txtDigito.Text, "0")

    If lOpcaoSelecionada = 1 Then
        txtPropostaSyas.Text = Format(txtPropostaSyas.Text, "00000000000000000000")
    End If

    If gM46V999.gfPreenchido(txtDV.Text) Then
        txtDV.Text = Format(gM46V999.CDblx(txtDV.Text), "000000")
    End If

    Call txtDataRecebimento_LostFocus

    '1. Consist�ncias para a pesquisa
    '1a. Efetuando consist�ncia no N�mero do PI

    If Not gM46V999.gfPreenchido(txtNumeroPI.Text) Then
        Call gM46V999.gpGraLog(1, "FNP0020 - N�mero do PI n�o preenchido.")

        If txtNumeroPI.Visible = True And txtNumeroPI.Enabled = True Then
            txtNumeroPI.SetFocus
        End If

        Exit Function
    ElseIf gM46V999.CDblx(txtNumeroPI.Text) = 0 Then
        Call gM46V999.gpGraLog(1, "FNP0020 - N�mero do PI n�o preenchido.")

        If txtNumeroPI.Visible = True And txtNumeroPI.Enabled = True Then
            txtNumeroPI.SetFocus
        End If

        Exit Function
    End If

    If Not gM46V999.gfPreenchido(txtDigito.Text) Then
        Call gM46V999.gpGraLog(1, "FNP0041 - D�gito do PI n�o preenchido.")

        If txtDigito.Visible = True And txtDigito.Enabled = True Then
            txtDigito.SetFocus
        End If

        Exit Function
    End If

    '2. Calcula d�gito.
    Set clsA99V001 = New clsA99V001
    Set clsA99V001.M46V999 = gM46V999

    If clsA99V001.DIG_YASUDA(txtNumeroPI.Text, lDigito, lMensagem) = False Then
        Exit Function
    End If

    Set clsA99V001 = Nothing

    If UCase(txtDigito.Text) <> lDigito Then
        'Ficha 1000449 - Incluir campo Referencia Marsh para digita��o
        If UCase(gM46V999.gNomUsuario6) = "TJGOME" Or UCase(gM46V999.gNomUsuario6) = "LVDELA" Or UCase(gM46V999.gNomUsuario6) = "SIMONY" Or UCase(gM46V999.gNomUsuario6) = "MFBARR" Then
            Call MsgBox(lDigito)
        Else
            Call gM46V999.gpGraLog(1, "FNP0041 - D�gito do PI inv�lido.")
        End If

        txtDigito.Text = ""

        If txtDigito.Visible = True And txtDigito.Enabled = True Then
            txtDigito.SetFocus
        End If

        Exit Function
    End If

    '1c. Efetuando consist�ncia na data de recebimento

    If Not gM46V999.gfPreenchido(txtDataRecebimento.Text) Then
        Call gM46V999.gpGraLog(1, "FGE0079 - Data de recebimento inv�lida(n�o preenchida).")

        If txtDataRecebimento.Visible = True And txtDataRecebimento.Enabled = True Then
            txtDataRecebimento.SetFocus
        End If

        Exit Function
    ElseIf Not IsDate(txtDataRecebimento.Text) Then
        Call gM46V999.gpGraLog(1, "FGE0079 - Data de recebimento inv�lida(" & txtDataRecebimento.Text & ").")

        txtDataRecebimento.Text = ""

        If txtDataRecebimento.Visible = True And txtDataRecebimento.Enabled = True Then
            txtDataRecebimento.SetFocus
        End If

        Exit Function
    ElseIf CVDate(txtDataRecebimento.Text) > Date And gM46V999.PerfilAcesso <> e_Perfil_DIG_Gestor Then
        Call gM46V999.gpGraLog(1, "FGE0079 - Data de recebimento inv�lida (superior a data de hoje:" & Date & ")")

        txtDataRecebimento.Text = ""

        If txtDataRecebimento.Visible = True And txtDataRecebimento.Enabled = True Then
            txtDataRecebimento.SetFocus
        End If

        Exit Function
    End If

    '1d. Efetuando consist�ncia no N�mero da proposta
    If lOpcaoSelecionada = 1 Then
        If Not gM46V999.gfPreenchido(txtPropostaSyas.Text) Then
            Call gM46V999.gpGraLog(1, "FNP0022 - N�mero da Proposta n�o preenchido.")

            If txtPropostaSyas.Visible = True And txtPropostaSyas.Enabled = True Then
                txtPropostaSyas.SetFocus
            End If

            Exit Function
        ElseIf gM46V999.CDblx(txtPropostaSyas.Text) = 0 Then
            Call gM46V999.gpGraLog(1, "FNP0022 - N�mero da Proposta n�o preenchido.")

            If txtPropostaSyas.Visible = True And txtPropostaSyas.Enabled = True Then
                txtPropostaSyas.SetFocus
            End If

            Exit Function
        End If                                                            'informado o n�mero da proposta.

        '1e. Calcula d�gito.
        Set clsA99V001 = New clsA99V001
        Set clsA99V001.M46V999 = gM46V999

        If clsA99V001.DIG_MODULO11(Left(txtPropostaSyas.Text, 19), "20191817161514131211100908070605040302", lDigitoNP, _
                                   lMensagem) = False Then
            Exit Function
        End If

        Set clsA99V001 = Nothing

        If Right(txtPropostaSyas.Text, 1) <> lDigitoNP Then
            Call gM46V999.gpGraLog(1, "FNL0041 - D�gito do n�mero da proposta est� inv�lido.")
            If txtPropostaSyas.Visible = True And txtPropostaSyas.Enabled = True Then
                txtPropostaSyas.SetFocus
            End If
            Exit Function
        End If

        '1f. Efetuando consist�ncia no D�gito Verificador
        If Not gM46V999.gfPreenchido(txtDV.Text) Then
            Call gM46V999.gpGraLog(1, "FNP0023 - D�gito Verificador n�o preenchido.")

            If txtDV.Visible = True And txtDV.Enabled = True Then
                txtDV.SetFocus
            End If

            Exit Function
        End If
    End If
    If lOpcaoSelecionada = 3 Then
        If Not gM46V999.gfPreenchido(txtMTVVD.Text) Then
            Call gM46V999.gpGraLog(1, "FNP0041 - Digite o MT da proposta VVD.")
            txtMTVVD.SetFocus
            Exit Function
        Else
            If Not IsNumeric(txtMTVVD.Text) Then
                Call gM46V999.gpGraLog(1, "FNP0041 - Proposta VVD inv�lida.")
                txtMTVVD.SetFocus
                Exit Function
            End If
            If Len(txtMTVVD.Text) < 6 Then
                Call gM46V999.gpGraLog(1, "FNP0041 - Proposta VVD inv�lida.")
                txtMTVVD.SetFocus
                Exit Function
            End If
        End If
    End If

    'Ficha 1000449 - Incluir campo Referencia Marsh para digita��o
    If lOpcaoSelecionada = 4 Then
        If Not gM46V999.gfPreenchido(txtMTVVD.Text) Then
            Call gM46V999.gpGraLog(1, "FNP0041 - Digite a Referencia Marsh.")
            txtMTVVD.SetFocus
            Exit Function
        Else
            If Not IsNumeric(txtMTVVD.Text) Then
                Call gM46V999.gpGraLog(1, "FNP0041 - a Referencia Marsh inv�lida.")
                txtMTVVD.SetFocus
                Exit Function
            End If
            If Len(txtMTVVD.Text) < 6 Then
                Call gM46V999.gpGraLog(1, "FNP0041 - Referencia Marsh inv�lida.")
                txtMTVVD.SetFocus
                Exit Function
            End If
        End If
    End If

    mfConsistirDados = True

End Function

Private Sub mpLimparCampos()
    txtPropostaSyas.Text = ""
    txtDigito.Text = ""
    txtDV.Text = ""
    mAntPI = 0
End Sub

Private Sub mpLimparErros()
    Dim lI As Long
    frmT46V012B.panAviso.Width = 10155
    frmT46V012B.flxResultado.Width = 9735
    With frmT46V012B.flxResultado
        .Cols = 2
        .ColWidth(0) = 500
        .ColWidth(1) = 9200
        .Rows = 3
        .TextMatrix(0, 1) = "Siscad"
        .TextMatrix(1, 1) = "Consist�ncias"
        .TextMatrix(2, 1) = "Pr�mio"
        For lI = 0 To 1
            .Row = lI
            .Col = 0
            Set .CellPicture = Nothing
        Next lI
        mQtdErros = 0
        mLinha = -1
    End With
End Sub

Private Sub txtPropostaSyas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If

End Sub

Private Function mfIncluirSiscad(pMensagem As String) As Boolean
    Dim lInsert As String
    Dim lAuxiliar As String
    Dim bdRamosDiversos As ADODB.Connection

    If gM46V999.gfAbrBasDados("RAMOSDIVERSOS", gM46V999.gTipDIVERSOS, gM46V999.gEndDIVERSOS, bdRamosDiversos, pMensagem) = False Then
        Exit Function
    End If
    lInsert = " INSERT INTO [P0044800].[dbo].[T0044800]([NOME_RAZAO], "
    lInsert = lInsert & " [NOM_ABREV], [NOM_CONTATO], [TIP_PESSOA], "
    lInsert = lInsert & " [NUM_CGC_CPF], [COD_TIP_LOGR], [NOM_LOGR], "
    lInsert = lInsert & " [DSC_COMPL_LOGR], [NOM_BAIRRO], [NOM_CID], "
    lInsert = lInsert & " [SIG_UF], [NUM_CEP], [NUM_DDD_FONE], [NUM_FONE], "
    lInsert = lInsert & " [NUM_DDD_FAX], [NUM_FAX], [DAT_NASC], [SEXO], "
    lInsert = lInsert & " [TIP_DOCUM], [NUM_DOCUM], [NOM_ORG_EXPED], [DAT_EXPED], "
    lInsert = lInsert & " [EMAIL], [NOM_PROFISSAO], [RAMO_ATIV_EMPR], "
    lInsert = lInsert & " [NOME_CONTATO], [COD_BANCO], [COD_AGEN], [NUM_CTA_COR], "
    lInsert = lInsert & " [DIG_CTA_COR], [COD_IDENTIFICADOR], [COD_TIP_LOGR_CORRESP], "
    lInsert = lInsert & " [NOM_LOGR_CORRESP], [DSC_COMPL_LOGR_CORRESP], "
    lInsert = lInsert & " [NOM_BAIRRO_CORRESP], [NOM_CID_CORRESP], [SIG_UF_CORRESP], "
    lInsert = lInsert & " [NUM_CEP_CORRESP], [TIP_CORRETOR], [TIP_SEGURADO], "
    lInsert = lInsert & " [TIP_ESTIPULANTE], [TIP_OFICINA], [TIP_VISTORIADOR], "
    lInsert = lInsert & " [TIP_BENEFICIARIO], [TIP_IRREGULAR], [DAT_CADASTRO], "
    lInsert = lInsert & " [DAT_ALTERACAO], [COD_USER], [DOC_ANEXOS], [NOM_LOCTRAB], "
    lInsert = lInsert & " [SUBTIP_PESSOA], [ALTERA_NOME], [NUM_PI], [DIVERGENCIA],[RENDAFATU])"
    lInsert = lInsert & "  VALUES("
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(mobjstcPed.Nom_Segurado) & ","
    lInsert = lInsert & "'',"
    lInsert = lInsert & "'',"
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(mobjstcPed.Tip_Pessoa) & ","
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(mobjstcPed.Num_Cgc_Cpf) & ","
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(mobjstcPed.objstcPedEndereco.Cod_Tip_Logr) & ","
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(mobjstcPed.objstcPedEndereco.Nom_Logr) & ","
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(mobjstcPed.objstcPedEndereco.Dsc_Compl_Logr) & ","
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(mobjstcPed.objstcPedEndereco.Nom_Bairro) & ","
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(mobjstcPed.objstcPedEndereco.Nom_Cid) & ","
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(mobjstcPed.objstcPedEndereco.Sig_UF) & ","
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(mobjstcPed.objstcPedEndereco.Num_CEP) & ","
    lInsert = lInsert & "0,"
    lAuxiliar = mfRetornarNumero(mobjstcPed.objstcPedEndereco.Num_Fone)
    lAuxiliar = val(Right(lAuxiliar, 8))
    lInsert = lInsert & lAuxiliar & ","
    lInsert = lInsert & "null,null, null,"
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(mobjstcPed.Sexo) & ","

    '[FRH] - 14/05/2009 - DE COD_TIP_DOC P/ COD_DOC
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(mobjstcPed.Cod_Doc) & ","

    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(mobjstcPed.Num_Doc) & ","
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(mobjstcPed.Org_Exp) & ","
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(mobjstcPed.Dat_Emis_Doc) & ","
    lInsert = lInsert & "null,"
    'Ficha 900417 - CAMPO NOM_PROFISSAO DA TABELA
    If gM46V999.gCodProduto >= 968 Then
        lInsert = lInsert & gM46V999.gfFormatarCamposSQL(Left(mobjstcPed.Nom_Empresa, 20)) & ","
    Else
        lInsert = lInsert & "null,"
    End If
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(" " & mobjstcPed.gLetraRamoAtiv) & ","
    lInsert = lInsert & "null,null,null,null,null, "
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(mobjstcPed.Cod_Segurado) & ","
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(mobjstcPed.Cod_Tip_Logr) & ","
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(mobjstcPed.Nom_Logr) & ","
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(mobjstcPed.Dsc_Compl_Logr) & ","
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(mobjstcPed.Nom_Bairro) & ","
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(mobjstcPed.Nom_Cid) & ","
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(mobjstcPed.Sig_UF) & ","
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(mobjstcPed.Num_CEP) & ","
    lInsert = lInsert & "null, null, null,null,null, null,null, "
    lInsert = lInsert & Format(Date, "yyyymmdd") & ","
    lInsert = lInsert & "0,"
    'Ficha 901659 - n�o utilizar o gM46V999.gNomUsuario pois est� dando erro por ser maior que o campo da tabela
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(gM46V999.gNomUsuario6) & ","
    lInsert = lInsert & "null,"
    'Ficha 900417 - CAMPO NOM_EMPRESA DA TABELA
    If gM46V999.gCodProduto >= 968 Then
        lInsert = lInsert & "null,"
    Else
        lInsert = lInsert & gM46V999.gfFormatarCamposSQL(Left(mobjstcPed.Nom_Empresa, 20)) & ","
    End If
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(mobjstcPed.Tip_Pessoa) & ","
    lInsert = lInsert & " 0, null,null,"
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(mobjstcPed.Val_RendaFatu) & ")"

    If gM46V999.gfExeSQL(bdRamosDiversos, lInsert, pMensagem) <> 0 Then
        Exit Function
    End If

    mfIncluirSiscad = True

End Function

Private Function mfAlterarSiscad(pMensagem As String) As Boolean
    Dim lUpdate As String
    Dim lAuxiliar As String
    Dim lAux As Long
    Dim bdRamosDiversos As ADODB.Connection

    If gM46V999.gfAbrBasDados("RAMOSDIVERSOS", gM46V999.gTipDIVERSOS, gM46V999.gEndDIVERSOS, bdRamosDiversos, pMensagem) = False Then
        Exit Function
    End If


    lUpdate = " Update [P0044800].[dbo].[T0044800]"
    'Ver como o nome ser� atualizado
    'Ficha 900417
    '    If mobjstcPed.Nom_Empresa <> "" Then
    '        lAux = InStr(1, mobjstcPed.Nom_Segurado, Left(mobjstcPed.Nom_Empresa, 10))
    '        lAuxiliar = mobjstcPed.Nom_Segurado
    '        If lAux > 0 Then
    '            Do While Mid(lAuxiliar, lAux, 1) <> " "
    '                lAuxiliar = Left(lAuxiliar, lAux - 1)
    '                lAux = lAux - 1
    '                If lAux = 0 Then GoTo Grava
    '            Loop
    '            lAuxiliar = Left(lAuxiliar, lAux - 1)
    '        End If
    'Grava:
    '        lUpdate = lUpdate & " set [NOME_RAZAO] =" & gM46V999.gfFormatarCamposSQL(lAuxiliar)
    '    Else
    '        lUpdate = lUpdate & " set [NOME_RAZAO] =" & gM46V999.gfFormatarCamposSQL(mobjstcPed.Nom_Segurado)
    '    End If

    lUpdate = lUpdate & " set [NOME_RAZAO] =" & gM46V999.gfFormatarCamposSQL(mobjstcPed.Nom_Segurado)

    If gM46V999.gfPreenchido(mobjstcPed.objstcPedEndereco.Num_Fone) Then
        lAuxiliar = mfRetornarNumero(mobjstcPed.objstcPedEndereco.Num_Fone)
        If Len(lAuxiliar) > 8 Then
            lUpdate = lUpdate & " ,[NUM_DDD_FONE] = " & Left(lAuxiliar, 2)
        End If
        lUpdate = lUpdate & " ,[NUM_FONE] = " & Right(lAuxiliar, 8)
    End If
    lUpdate = lUpdate & " ,[TIP_PESSOA] = " & gM46V999.gfFormatarCamposSQL(mobjstcPed.Tip_Pessoa)
    lUpdate = lUpdate & " ,[COD_TIP_LOGR] = " & gM46V999.gfFormatarCamposSQL(mobjstcPed.objstcPedEndereco.Cod_Tip_Logr)
    lUpdate = lUpdate & " ,[NOM_LOGR]= " & gM46V999.gfFormatarCamposSQL(mobjstcPed.objstcPedEndereco.Nom_Logr)
    lUpdate = lUpdate & " ,[DSC_COMPL_LOGR] =" & gM46V999.gfFormatarCamposSQL(mobjstcPed.objstcPedEndereco.Dsc_Compl_Logr)
    lUpdate = lUpdate & " ,[NOM_BAIRRO] =" & gM46V999.gfFormatarCamposSQL(mobjstcPed.objstcPedEndereco.Nom_Bairro)
    lUpdate = lUpdate & " ,[NOM_CID]= " & gM46V999.gfFormatarCamposSQL(mobjstcPed.objstcPedEndereco.Nom_Cid)
    lUpdate = lUpdate & " ,[SIG_UF] =" & gM46V999.gfFormatarCamposSQL(mobjstcPed.objstcPedEndereco.Sig_UF)
    lUpdate = lUpdate & " ,[NUM_CEP]= " & gM46V999.gfFormatarCamposSQL(mobjstcPed.objstcPedEndereco.Num_CEP)
    lUpdate = lUpdate & " ,[SEXO]= " & gM46V999.gfFormatarCamposSQL(mobjstcPed.Sexo)

    '[FRH] - 14/05/2009 - DE COD_TIP_DOC P/ COD_DOC
    lUpdate = lUpdate & " ,[TIP_DOCUM]=" & gM46V999.gfFormatarCamposSQL(mobjstcPed.Cod_Doc)

    lUpdate = lUpdate & " ,[NUM_DOCUM]=" & gM46V999.gfFormatarCamposSQL(mobjstcPed.Num_Doc)
    lUpdate = lUpdate & " ,[NOM_ORG_EXPED]=" & gM46V999.gfFormatarCamposSQL(mobjstcPed.Org_Exp)
    lUpdate = lUpdate & " ,[DAT_EXPED]= " & gM46V999.gfFormatarCamposSQL(mobjstcPed.Dat_Emis_Doc)
    lUpdate = lUpdate & " ,[RAMO_ATIV_EMPR]= " & gM46V999.gfFormatarCamposSQL(" " & mobjstcPed.gLetraRamoAtiv)
    lUpdate = lUpdate & " ,[COD_IDENTIFICADOR]=" & gM46V999.gfFormatarCamposSQL(mobjstcPed.Cod_Segurado)
    lUpdate = lUpdate & " ,[COD_TIP_LOGR_CORRESP]= " & gM46V999.gfFormatarCamposSQL(mobjstcPed.Cod_Tip_Logr)
    lUpdate = lUpdate & " ,[NOM_LOGR_CORRESP]=" & gM46V999.gfFormatarCamposSQL(mobjstcPed.Nom_Logr)
    lUpdate = lUpdate & " ,[DSC_COMPL_LOGR_CORRESP]= " & gM46V999.gfFormatarCamposSQL(mobjstcPed.Dsc_Compl_Logr)
    lUpdate = lUpdate & " ,[NOM_BAIRRO_CORRESP]=" & gM46V999.gfFormatarCamposSQL(mobjstcPed.Nom_Bairro)
    lUpdate = lUpdate & " ,[NOM_CID_CORRESP]=" & gM46V999.gfFormatarCamposSQL(mobjstcPed.Nom_Cid)
    lUpdate = lUpdate & " ,[SIG_UF_CORRESP]= " & gM46V999.gfFormatarCamposSQL(mobjstcPed.Sig_UF)
    lUpdate = lUpdate & " ,[NUM_CEP_CORRESP]=" & gM46V999.gfFormatarCamposSQL(mobjstcPed.Num_CEP)
    lUpdate = lUpdate & " ,[DAT_ALTERACAO]=" & Format(Date, "yyyymmdd")
    'Ficha 901659 - n�o utilizar o gM46V999.gNomUsuario pois est� dando erro por ser maior que o campo da tabela
    lUpdate = lUpdate & " ,[COD_USER]=" & gM46V999.gfFormatarCamposSQL(gM46V999.gNomUsuario6)
    'Ficha 900417
    If gM46V999.gCodProduto >= 968 Then
        lUpdate = lUpdate & " ,[NOM_PROFISSAO]=" & gM46V999.gfFormatarCamposSQL(Left(mobjstcPed.Nom_Empresa, 20))
    Else
        lUpdate = lUpdate & " ,[NOM_LOCTRAB]=" & gM46V999.gfFormatarCamposSQL(Left(mobjstcPed.Nom_Empresa, 20))
    End If
    lUpdate = lUpdate & " ,[SUBTIP_PESSOA]=" & gM46V999.gfFormatarCamposSQL(mobjstcPed.Tip_Pessoa)
    lUpdate = lUpdate & " ,[RENDAFATU]=" & gM46V999.gfFormatarCamposSQL(mobjstcPed.Val_RendaFatu)
    lUpdate = lUpdate & "  where [NUM_CGC_CPF] =" & gM46V999.gfFormatarCamposSQL(mobjstcPed.Num_Cgc_Cpf)

    If gM46V999.gfExeSQL(bdRamosDiversos, lUpdate, pMensagem) <> 0 Then
        Exit Function
    End If

    mfAlterarSiscad = True

End Function

Private Function mfGravarEndCobranca(pMensagem As String) As Boolean
'Carregar os endere�o de cobran�a
    Dim lInsert As String
    Dim bdRamosDiversos As ADODB.Connection

    If gM46V999.gfAbrBasDados("RAMOSDIVERSOS", gM46V999.gTipDIVERSOS, gM46V999.gEndDIVERSOS, bdRamosDiversos, pMensagem) = False Then
        Exit Function
    End If

    lInsert = "INSERT INTO [P0044800].[dbo].[T0044807]"
    lInsert = lInsert & "([num_pi], [num_cgc_cpf], [dat_emis], [dat_ter_vig], "
    lInsert = lInsert & "[num_apol], [num_endo], [cod_tip_logr], [nom_logr], "
    lInsert = lInsert & "[dsc_compl_logr], [nom_bairro], [nom_cid], [sig_uf], "
    lInsert = lInsert & "[num_cep], [num_ddd_fone], [num_fone], [dat_cadastro], "
    lInsert = lInsert & "[dat_alteracao], [cod_user]) VALUES("
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(val(txtNumeroPI.Text)) & ","
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(mobjstcPed.Num_Cgc_Cpf) & ","
    lInsert = lInsert & "0,"
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(mobjstcPed.Dat_Ter_Vig) & ","
    lInsert = lInsert & "0,0,"
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(mobjstcPed.Cod_Tip_Logr) & ","
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(mobjstcPed.Nom_Logr) & ","
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(mobjstcPed.Dsc_Compl_Logr) & ","
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(mobjstcPed.Nom_Bairro) & ","
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(mobjstcPed.Nom_Cid) & ","
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(mobjstcPed.Sig_UF) & ","
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(mobjstcPed.Num_CEP) & ","
    '1001046 - Telefone do Segurado 8 Digitos
    lInsert = lInsert & "null, " & gM46V999.gfFormatarCamposSQL(mobjstcPed.Num_Fone) & ", "
    lInsert = lInsert & Format(Date, "yyyymmdd") & ","
    lInsert = lInsert & Format(Date, "yyyymmdd") & ","
    'Ficha 901659 - n�o utilizar o gM46V999.gNomUsuario pois est� dando erro por ser maior que o campo da tabela
    lInsert = lInsert & gM46V999.gfFormatarCamposSQL(gM46V999.gNomUsuario6) & ")"

    If gM46V999.gfExeSQL(bdRamosDiversos, lInsert, pMensagem) <> 0 Then
        If Not pMensagem Like "*Cannot insert duplicate key*" Then
            Exit Function
        End If
    End If

    mfGravarEndCobranca = True
End Function

Private Sub mnuExiRelatorio_Click()
    Dim lAppExcel As excel.Application    'Aplica��o Excel.
    Dim lwksPrincipal As excel.Worksheet    'Planilha excel.
    Dim bdRamosDiversos As ADODB.Connection
    Dim lArquivo As String
    Dim lLinha As Long
    Dim lI As Integer
    Dim lAuxiliar As String
    Dim lData As String
    Dim pMensagem As String

    If gM46V999.gfAbrBasDados("RAMOSDIVERSOS", gM46V999.gTipDIVERSOS, gM46V999.gEndDIVERSOS, bdRamosDiversos, pMensagem) = False Then
        Exit Sub
    End If

    Set lobjTriagem = New clsA46v729A
    Set lobjTriagem.M46V999 = gM46V999

    Set lobjstcTriagem = New stcA46v729B
    Set lobjTriagem.ConexaoBD = bdRamosDiversos

    lData = InputBox("Data de pesquisa de triagem", "Relat�rio de triagem", Date)

    If IsDate(lData) = False Then
        Call gM46V999.gpGraLog(1, "Data inv�lida!")
        Exit Sub
    End If

    lobjstcTriagem.Dat_Triagem = gM46V999.gfForData(1, lData)

    If Not lobjTriagem.gfCarregarColecao(lobjstcTriagem) Then
        Call gM46V999.gpGraLog(1, lobjTriagem.UltimoErro)
        Exit Sub
    End If

    Set lAppExcel = CreateObject("Excel.Application")

    lAppExcel.Workbooks.Add
    lLinha = 1

    Set lwksPrincipal = lAppExcel.ActiveSheet

    lwksPrincipal.Cells(lLinha, 1) = "Num. PI"
    lwksPrincipal.Cells(lLinha, 2) = "Num Syas"
    lwksPrincipal.Cells(lLinha, 3) = "Data"
    lwksPrincipal.Cells(lLinha, 4) = "Hora"
    lwksPrincipal.Cells(lLinha, 5) = "Usu�rio"
    lwksPrincipal.Cells(lLinha, 6) = "Status"

    For Each lobjstcTriagem In lobjTriagem.Colecao
        lLinha = lLinha + 1
        lwksPrincipal.Cells(lLinha, 1) = CStr(lobjstcTriagem.NUM_PI)
        lwksPrincipal.Cells(lLinha, 2) = CStr("'" & lobjstcTriagem.NOSSO_NUMERO)
        lwksPrincipal.Cells(lLinha, 3) = CStr("'" & Format(Mid(lobjstcTriagem.Dat_Triagem, 7, 2) & "/" & Mid(lobjstcTriagem.Dat_Triagem, 5, 2) & "/" & Left(lobjstcTriagem.Dat_Triagem, 4), "dd/mm/yyyy"))
        lwksPrincipal.Cells(lLinha, 4) = CStr(Mid(lobjstcTriagem.Dat_Triagem, 9, 2) & ":" & Mid(lobjstcTriagem.Dat_Triagem, 11, 2) & ":" & Mid(lobjstcTriagem.Dat_Triagem, 13, 2))
        lwksPrincipal.Cells(lLinha, 5) = CStr(lobjstcTriagem.User_Triagem)
        Select Case lobjstcTriagem.STATUS
        Case e_status_triagem_Erro, e_status_triagem_ErroTriagem
            lAuxiliar = "Diverg�ncia/Suspenso"
        Case e_status_triagem_ErroPremio
            lAuxiliar = "Pr�mio e/ou parcela divergente/Suspenso"
        Case e_status_triagem_ErroSiscad
            lAuxiliar = "Validade SISCAD/Suspenso"
        Case e_status_triagem_NDefinido
            lAuxiliar = "N�o emitido e nem suspenso"
        Case e_status_triagem_NLoc
            lAuxiliar = "N�o localizada"
        Case e_status_triagem_OK
            lAuxiliar = "Emitido"
        End Select
        lwksPrincipal.Cells(lLinha, 6) = lAuxiliar
    Next lobjstcTriagem

    For lI = 1 To 6
        lwksPrincipal.Columns(lI).EntireColumn.AutoFit
    Next lI

    With lwksPrincipal.Range("A1:F" & lLinha)
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeLeft).ColorIndex = xlAutomatic
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeTop).ColorIndex = xlAutomatic
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeBottom).ColorIndex = xlAutomatic
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Weight = xlThin
        .Borders(xlEdgeRight).ColorIndex = xlAutomatic
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideVertical).Weight = xlThin
        .Borders(xlInsideVertical).ColorIndex = xlAutomatic
        On Error Resume Next
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).Weight = xlThin
        .Borders(xlInsideHorizontal).ColorIndex = xlAutomatic
    End With

    lwksPrincipal.Range("A1:F1").Interior.ColorIndex = 15

    lAppExcel.Visible = True

    Exit Sub

ControleErro:

    Call gM46V999.gpGraLog(1, Err.Description)

End Sub
Private Sub mpObtProduto(ByVal pCodProduto As Integer, ByVal pCodRamo As Long)
'Procedure: obt�m produto e informa��es dependentes de produto.

'Par�metros de entrada...pCodProduto.....c�digo do produto.
'                        pCodProtocolo...c�digo do protocolo.
'                        pCodSequencia...c�digo da seq��ncia.

    Dim lbdTAB_PROD As ADODB.Connection    'Base de dados - TAB_PROD.
    Dim lCodProtocolo As Long           'C�digo do protocolo.
    Dim lCodSequencia As Long           'C�digo da seq��ncia.
    Dim lI As Byte    'Utilizado com For...Next.
    Dim lMensagem As String     'Mensagem.
    Dim lPosicao As Integer   'Utilizado com InStr.
    Dim lrsTAB_PROD As clsYasRecordSet    'Registro - TAB_PROD.
    Dim lCarga As Boolean
    Dim lSelect As String

    '1. Inicializa vari�veis globais - dados do produto.

    gM46O001.gProTipProtocolo = 1
    gM46O001.gProNome = " "
    gM46O001.gProLimComissao = 0
    gM46O001.gProComBasica = 0
    gM46V999.gProCusApolice = 0
    gM46V999.gProPreMinParcela = 0
    gM46V999.gProPremioMinimo = 0
    gM46V999.gProPreMinEndosso = 0
    gM46O001.gProLimMarComercial = 0
    gM46V999.gProQuaMaxParcelas = 0
    'simonyJuros
    'gM46O001.gProJurMinAnt01 = 3
    'gM46O001.gProJurMinAnt02 = 3
    'gM46O001.gProJurMinAnt03 = 3
    'gM46O001.gProJurMinAnt04 = 3
    'gM46O001.gProJurMinAnt05 = 3
    'gM46O001.gProJurMinAnt06 = 3
    'gM46O001.gProJurMinAnt07 = 3
    'gM46O001.gProJurMinAnt08 = 3
    'gM46O001.gProJurMinAnt09 = 3
    'gM46O001.gProJurMinAnt10 = 3
    'gM46O001.gProJurMinAnt11 = 3
    'gM46O001.gProJurMinAnt12 = 3
    'gM46O001.gProJurMinPos01 = 3
    'gM46O001.gProJurMinPos02 = 3
    'gM46O001.gProJurMinPos03 = 3
    'gM46O001.gProJurMinPos04 = 3
    'gM46O001.gProJurMinPos05 = 3
    'gM46O001.gProJurMinPos06 = 3
    'gM46O001.gProJurMinPos07 = 3
    'gM46O001.gProJurMinPos08 = 3
    'gM46O001.gProJurMinPos09 = 3
    'gM46O001.gProJurMinPos10 = 3
    'gM46O001.gProJurMinPos11 = 3
    'gM46O001.gProJurMinPos12 = 3
    gM46O001.gProJurMinPadrao = 3
    gM46V111.gProJurIOF = 0

    gM46V999.gCodProduto = 0
    gM46V999.gCodProtocolo = 0
    gM46V999.gCodSequencia = 0


    '2. Se produto zerado, encerra processo.
    If pCodProduto = 0 Or pCodRamo = 0 Then
        Exit Sub
    End If

    '3. Abre arquivo.
    If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTAB_PROD, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        gM46V999.gCodProduto = 0
        gM46V999.gCodProtocolo = 0
        gM46V999.gCodSequencia = 0
        Exit Sub
    End If

    '4. Obt�m produto informado.
    lSelect = "SELECT * FROM TAB_PRODUTO WHERE Cod_Produto = " & pCodProduto & " AND Cod_Ramo = " & _
              pCodRamo
    If gM46V999.gfObtRegistro(lbdTAB_PROD, lSelect, lrsTAB_PROD, lMensagem) = False Then
        Call gM46V999.gpFecha2(lbdTAB_PROD)
        Call gM46V999.gpGraLog(2, lMensagem)
        gM46V999.gCodProduto = 0
        gM46V999.gCodProtocolo = 0
        gM46V999.gCodSequencia = 0
        Exit Sub
    End If
    If lrsTAB_PROD.EOF = True Then
        Call gM46V999.gpFecha2(lbdTAB_PROD)
        Call gM46V999.gpGraLog(1, "FNL0002 - " & gM46V999.gPrefixo & "Produto n�o localizado (" & _
                                  Format$(pCodProduto, "000") & "/" & Format$(pCodRamo, "000") & ").")
        gM46V999.gCodProduto = 0
        gM46V999.gCodProtocolo = 0
        gM46V999.gCodSequencia = 0
        Exit Sub
    End If

    '5. Carrega vari�veis globais.
    gM46V999.gCodProduto = lrsTAB_PROD!Cod_Produto
    gM46V999.gCodProtocolo = lrsTAB_PROD!Cod_Protocolo
    gM46V999.gCodSequencia = lrsTAB_PROD!Cod_Sequencia
    gM46O001.gProTipProtocolo = lrsTAB_PROD!Tip_Prot
    gM46O001.gProNome = lrsTAB_PROD!Nom_Prot
    gM46O001.gProLimComissao = lrsTAB_PROD!Per_Cms_Limite
    gM46O001.gProComBasica = lrsTAB_PROD!Per_Cms_Basica
    gM46V999.gProCusEndosso = lrsTAB_PROD!Val_Cus_Emis_Endo
    gM46V999.gProCusApolice = lrsTAB_PROD!Val_Cus_Emis_Apol
    gM46V999.gProPreMinParcela = lrsTAB_PROD!Par_Min_Parcela
    gM46V999.gProPreMinParcelaEndosso = lrsTAB_PROD!Par_Min_Endosso
    gM46V999.gProPremioMinimo = lrsTAB_PROD!Pre_Min_Parcela
    gM46V999.gProPreMinEndosso = lrsTAB_PROD!Pre_Min_Endosso
    gM46O001.gProLimMarComercial = lrsTAB_PROD!Lim_Marg_Coml
    gM46V999.gProQuaMaxParcelas = lrsTAB_PROD!Qua_Max_Parcelas

    'tbmjuros
    'SimonyJuros

    'gM46O001.gProJurMinAnt01 = lrsTAB_PROD!Jur_Min_Ant_01
    'gM46O001.gProJurMinAnt02 = lrsTAB_PROD!Jur_Min_Ant_02
    'gM46O001.gProJurMinAnt03 = lrsTAB_PROD!Jur_Min_Ant_03
    'gM46O001.gProJurMinAnt04 = lrsTAB_PROD!Jur_Min_Ant_04
    'gM46O001.gProJurMinAnt05 = lrsTAB_PROD!Jur_Min_Ant_05
    'gM46O001.gProJurMinAnt06 = lrsTAB_PROD!Jur_Min_Ant_06
    'gM46O001.gProJurMinAnt07 = lrsTAB_PROD!Jur_Min_Ant_07
    'gM46O001.gProJurMinAnt08 = lrsTAB_PROD!Jur_Min_Ant_08
    'gM46O001.gProJurMinAnt09 = lrsTAB_PROD!Jur_Min_Ant_09
    'gM46O001.gProJurMinAnt10 = lrsTAB_PROD!Jur_Min_Ant_10
    'gM46O001.gProJurMinAnt11 = lrsTAB_PROD!Jur_Min_Ant_11
    'gM46O001.gProJurMinAnt12 = lrsTAB_PROD!Jur_Min_Ant_12
    'gM46O001.gProJurMinPos01 = lrsTAB_PROD!Jur_Min_Pos_01
    'gM46O001.gProJurMinPos02 = lrsTAB_PROD!Jur_Min_Pos_02
    'gM46O001.gProJurMinPos03 = lrsTAB_PROD!Jur_Min_Pos_03
    'gM46O001.gProJurMinPos04 = lrsTAB_PROD!Jur_Min_Pos_04
    'gM46O001.gProJurMinPos05 = lrsTAB_PROD!Jur_Min_Pos_05
    'gM46O001.gProJurMinPos06 = lrsTAB_PROD!Jur_Min_Pos_06
    'gM46O001.gProJurMinPos07 = lrsTAB_PROD!Jur_Min_Pos_07
    'gM46O001.gProJurMinPos08 = lrsTAB_PROD!Jur_Min_Pos_08
    'gM46O001.gProJurMinPos09 = lrsTAB_PROD!Jur_Min_Pos_09
    'gM46O001.gProJurMinPos10 = lrsTAB_PROD!Jur_Min_Pos_10
    'gM46O001.gProJurMinPos11 = lrsTAB_PROD!Jur_Min_Pos_11
    'gM46O001.gProJurMinPos12 = lrsTAB_PROD!Jur_Min_Pos_12

    gM46O001.gProJurMinPadrao = lrsTAB_PROD!Jur_Padrao
    gM46V111.gProJurIOF = lrsTAB_PROD!Jur_IOF
    Call gM46V999.gpFecha2(lbdTAB_PROD)

End Sub


Private Sub txtPropostaSyas_KeyUp(KeyCode As Integer, Shift As Integer)
    If Len(txtPropostaSyas.Text) = 20 Then
        SendKeys vbTab
    End If
End Sub

Function mfVerDizer(ByRef pobjstcPed As stcA46V702B, ByRef pMensagem As String) As Boolean
    Dim objPedido As clsA46V702A    'Objeto da classe clsA46V702A  (TAB_PED)
    Dim objTabVenc As clsA46V713A
    Dim lbdBasDados As ADODB.Connection    'Base de dados.
    Dim lrsTabela As clsYasRecordSet    'Registro da tabela.
    Dim lobjstcTextoEspec As stcA46V726B   'Objeto da classe de estrutura stcA46V726B (TAB_PED_ESPEC).
    Dim ObjTextoEspec As clsA46V726A    'Objeto da classe de estrutura clsA46V726A (TAB_PED_ESPEC).
    Dim lObjstcVenc As stcA46V713B    'Objeto da classe de estrutura stcA46V713B (TAB_PED_Venc)
    Dim lSql As String
    Dim lI As Integer
    Dim lNum_pi As Long
    Dim lNosNumero As String
    Dim codDizer As Integer
    Dim ForPagamento As Integer
    Dim TipoPagamento As Integer
    Dim lArrayDatVenc As Variant
    Dim lDataIrregular As Boolean

    On Error GoTo mfVerDizer_Err

    mfVerDizer = False

    lNum_pi = pobjstcPed.NUM_PI
    lNosNumero = pobjstcPed.NOSSO_NUMERO

    If val(lNum_pi) = 0 Then
        mfVerDizer = True
        Exit Function
    End If

    If gM46V999.gfAbrBasDados("GED", gM46V999.gTipGed, gM46V999.gEndGed, lbdBasDados, pMensagem) = False Then Exit Function

    lSql = "Select * "
    lSql = lSql & " from tab_triagem  where num_pi = " & lNum_pi

    If gM46V999.gfObtRegistro(lbdBasDados, lSql, lrsTabela, pMensagem) = False Then
        Call gM46V999.gpFecha2(lbdBasDados)
        Exit Function
    End If

    If lrsTabela.EOF Then
        mfVerDizer = True
        GoTo Finalizar
    End If

    If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then
        If lrsTabela("ind_relbens") = 2 Then
            'Procurar se h� dizer na TAB_PED_ESPEC com o c�digo 8401 e exclui
            codDizer = 8401
            If Not mfDeletaDizer(mobjstcPed, codDizer, pMensagem) Then
                Exit Function
            End If
        End If
        If lrsTabela("ind_relbens") = 1 Then
            'Procurar dizer na TAB_PED_DIZER_PAD com o c�digo 8401
            codDizer = 8401
            If Not mfProcuraDizer(mobjstcPed, codDizer, pMensagem) Then
                Exit Function
            End If
        End If

        If lrsTabela("ind_segpredio") = 2 Then
            'Procurar se h� dizer na TAB_PED_ESPEC com o c�digo 8565 e exclui
            codDizer = 8565
            If Not mfDeletaDizer(mobjstcPed, codDizer, pMensagem) Then
                Exit Function
            End If
        End If
        If lrsTabela("ind_segpredio") = 1 Then
            'Procurar dizer na TAB_PED_DIZER_PAD com o c�digo 8565
            codDizer = 8565
            If Not mfProcuraDizer(mobjstcPed, codDizer, pMensagem) Then
                Exit Function
            End If
        End If

        If lrsTabela("ind_segconteudo") = 2 Then
            'Procurar se h� dizer na TAB_PED_ESPEC com o c�digo 8621 e exclui
            codDizer = 8621
            If Not mfDeletaDizer(mobjstcPed, codDizer, pMensagem) Then
                Exit Function
            End If
        End If
        If lrsTabela("ind_segconteudo") = 1 Then
            'Procurar dizer na TAB_PED_DIZER_PAD com o c�digo 8621
            codDizer = 8621
            If Not mfProcuraDizer(mobjstcPed, codDizer, pMensagem) Then
                Exit Function
            End If
        End If

        '######### Antecipado #########
        If pobjstcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec = 1 Then
            If pobjstcPed.objstcCtrlEmis.Num_Bloq <> 0 Then    'carn�
                ForPagamento = 1    'carn�
                TipoPagamento = 1    'antecipado
            Else
                ForPagamento = 2    'd�bito
                TipoPagamento = 1    'antecipado
            End If
        Else
            '######### A 30 dias #########
            If pobjstcPed.Cod_Agen_Bas_Cor <> 0 Then
                ForPagamento = 2    'd�bito
                TipoPagamento = 2    'a 30 dias
            Else
                ForPagamento = 1    'carn�
                TipoPagamento = 2    'a 30 dias
            End If
        End If

        'Exibe a tela de ajuste de vencimento das parcelas.
        gProChamado = "T46V121A"
        Screen.MousePointer = vbDefault

        If TipoPagamento = 1 Then    'Antecipada
            If IsNull(lrsTabela("dat_venci_demais")) Then
                mfVerDizer = True
                GoTo Finalizar
            End If
            Call gpVerificarVencimento(pobjstcPed.Qtd_Parc, TipoPagamento, _
                                       lrsTabela("dat_venci_demais"), ForPagamento, pobjstcPed.Dat_Ini_Vig, _
                                       lArrayDatVenc, lDataIrregular)
        Else    'a 30 dias
            If IsNull(lrsTabela("dat_venci_primeira")) Then
                mfVerDizer = True
                GoTo Finalizar
            End If
            Call gpVerificarVencimento(pobjstcPed.Qtd_Parc, TipoPagamento, _
                                       lrsTabela("dat_venci_primeira"), ForPagamento, pobjstcPed.Dat_Ini_Vig, _
                                       lArrayDatVenc, lDataIrregular)

        End If

        If TipoPagamento = 1 Then
            Call frmT46V121A.gpInstanciarForm2(pobjstcPed.Qtd_Parc, True, lArrayDatVenc)
        Else
            Call frmT46V121A.gpInstanciarForm2(pobjstcPed.Qtd_Parc, False, lArrayDatVenc)
        End If

        Screen.MousePointer = vbHourglass
        'Salvar na cole��o

        'Salvo as datas de vencimento ajustadas
        Set objTabVenc = New clsA46V713A
        Set objTabVenc.M46V999 = gM46V999

        Set lObjstcVenc = New stcA46V713B
        Set lObjstcVenc.M46V999 = gM46V999

        lObjstcVenc.NOSSO_NUMERO = pobjstcPed.NOSSO_NUMERO
        If Not objTabVenc.mfExcluir(lObjstcVenc, "", "", "") Then
            Call mpInformarErro(objTabVenc.UltimoErro, 2, True)
        Else
            Set pobjstcPed.ColecaoPedVenc = New Collection
            For lI = 0 To (pobjstcPed.Qtd_Parc - 1)
                If lI < 11 Then
                    If IsDate(frmT46V121A.txtDtVenc(lI + 1).Text) = False Then
                        pMensagem = "FGE0082 - Data de vencimento da parcela " & lI + 1 & " inv�lida (n�o preenchida)."
                        Call mpInformarErro(pMensagem, 2, True)
                    Else
                        Set lObjstcVenc = New stcA46V713B
                        Set lObjstcVenc.M46V999 = gM46V999

                        lObjstcVenc.NOSSO_NUMERO = pobjstcPed.NOSSO_NUMERO
                        lObjstcVenc.Cod_Ramo = pobjstcPed.Cod_Ramo
                        lObjstcVenc.Num_Ped = pobjstcPed.Num_Ped
                        lObjstcVenc.Num_Parc = lI + 1
                        lObjstcVenc.DAT_VENC = gM46V999.gfForData(1, frmT46V121A.txtDtVenc(lI + 1).Text)
                        If Not objTabVenc.mfIncluir(lObjstcVenc, "", "", "") Then
                            pMensagem = objTabVenc.UltimoErro
                            Call mpInformarErro(pMensagem, 2, True)
                        Else
                            pobjstcPed.ColecaoPedVenc.Add lObjstcVenc, lObjstcVenc.Class_ID
                        End If
                    End If
                End If
            Next
        End If
        Unload frmT46V121A
    End If

    mfVerDizer = True

Finalizar:

    Exit Function

mfVerDizer_Err:
    pMensagem = "Verificar dados de Forma de Pagamento/Datas de Parcelas na guia de Triagem"
End Function


Function mfProcuraDizer(ByRef pobjstcPed As stcA46V702B, lCodDizer As Integer, ByRef pMensagem As String) As Boolean
    Dim objPedido As clsA46V702A    'Objeto da classe clsA46V702A  (TAB_PED)
    Dim objTabVenc As clsA46V713A
    Dim lbdTab_Ped_Dizer_Pad As ADODB.Connection
    Dim lrsTab_Ped_Dizer_Pad As clsYasRecordSet
    Dim lbdBasDados As ADODB.Connection    'Base de dados.
    Dim lrsTabela As clsYasRecordSet    'Registro da tabela.
    Dim lbdTab_Ped_Espec As ADODB.Connection
    Dim lrsTab_Ped_Espec As clsYasRecordSet
    Dim lobjstcTextoEspec As stcA46V726B   'Objeto da classe de estrutura stcA46V726B (TAB_PED_ESPEC).
    Dim ObjTextoEspec As clsA46V726A    'Objeto da classe de estrutura clsA46V726A (TAB_PED_ESPEC).
    Dim lObjstcVenc As stcA46V713B    'Objeto da classe de estrutura stcA46V713B (TAB_PED_Venc)
    Dim lSql As String
    Dim lI As Integer
    Dim lNum_Pag As Integer
    Dim lNum_pi As Long
    Dim lNosNumero As String
    Dim textoEspec1 As String
    Dim textoEspec2 As String
    Dim textoEspec3 As String
    Dim textoEspec4 As String
    Dim textoEspec5 As String
    Dim textoEspec6 As String
    Dim textoEspec7 As String
    Dim textoEspec8 As String
    Dim textoEspec9 As String
    Dim textoEspec10 As String
    Dim textoEspec11 As String
    Dim textoEspec12 As String
    Dim textoEspec13 As String
    Dim textoEspec14 As String
    Dim textoEspec15 As String
    Dim textoEspec16 As String
    Dim textoEspec17 As String
    Dim textoEspec18 As String
    Dim textoEspec19 As String
    Dim textoEspec20 As String
    Dim lSelect As String

    On Error GoTo mfProcuraDizer_Err

    mfProcuraDizer = False

    If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTab_Ped_Dizer_Pad, pMensagem) = False Then
        Call gM46V999.gpGraLog(2, pMensagem)
        Exit Function
    End If
    lSelect = "SELECT * FROM TAB_PED_DIZER_PAD WHERE COD_DIZER_PAD = " & lCodDizer
    lSelect = lSelect & " order by COD_DIZER_PAD , num_seq "

    If gM46V999.gfObtRegistro(lbdTab_Ped_Dizer_Pad, lSelect, lrsTab_Ped_Dizer_Pad, pMensagem) = False Then
        Call gM46V999.gpFecha2(lbdTab_Ped_Dizer_Pad)
        Call gM46V999.gpGraLog(2, pMensagem)
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    If lrsTab_Ped_Dizer_Pad.EOF Then
        MsgBox "Mensagem Padr�o " & lCodDizer & " n�o encontrada!" & vbCrLf & "Favor efetuar cadastro do Dizer para prosseguir emiss�o!", vbInformation
        Exit Function
    Else
        Do While Not lrsTab_Ped_Dizer_Pad.EOF
            Select Case lrsTab_Ped_Dizer_Pad("num_seq")
            Case 1
                textoEspec1 = Trim(lrsTab_Ped_Dizer_Pad("Dsc_Dizer_Endo"))
            Case 2
                textoEspec2 = Trim(lrsTab_Ped_Dizer_Pad("Dsc_Dizer_Endo"))
            Case 3
                textoEspec3 = Trim(lrsTab_Ped_Dizer_Pad("Dsc_Dizer_Endo"))
            Case 4
                textoEspec4 = Trim(lrsTab_Ped_Dizer_Pad("Dsc_Dizer_Endo"))
            Case 5
                textoEspec5 = Trim(lrsTab_Ped_Dizer_Pad("Dsc_Dizer_Endo"))
            Case 6
                textoEspec6 = Trim(lrsTab_Ped_Dizer_Pad("Dsc_Dizer_Endo"))
            Case 7
                textoEspec7 = Trim(lrsTab_Ped_Dizer_Pad("Dsc_Dizer_Endo"))
            Case 8
                textoEspec8 = Trim(lrsTab_Ped_Dizer_Pad("Dsc_Dizer_Endo"))
            Case 9
                textoEspec9 = Trim(lrsTab_Ped_Dizer_Pad("Dsc_Dizer_Endo"))
            Case 10
                textoEspec10 = Trim(lrsTab_Ped_Dizer_Pad("Dsc_Dizer_Endo"))
            Case 11
                textoEspec11 = Trim(lrsTab_Ped_Dizer_Pad("Dsc_Dizer_Endo"))
            Case 12
                textoEspec12 = Trim(lrsTab_Ped_Dizer_Pad("Dsc_Dizer_Endo"))
            Case 13
                textoEspec13 = Trim(lrsTab_Ped_Dizer_Pad("Dsc_Dizer_Endo"))
            Case 14
                textoEspec14 = Trim(lrsTab_Ped_Dizer_Pad("Dsc_Dizer_Endo"))
            Case 15
                textoEspec15 = Trim(lrsTab_Ped_Dizer_Pad("Dsc_Dizer_Endo"))
            Case 16
                textoEspec16 = Trim(lrsTab_Ped_Dizer_Pad("Dsc_Dizer_Endo"))
            Case 17
                textoEspec17 = Trim(lrsTab_Ped_Dizer_Pad("Dsc_Dizer_Endo"))
            Case 18
                textoEspec18 = Trim(lrsTab_Ped_Dizer_Pad("Dsc_Dizer_Endo"))
            Case 19
                textoEspec19 = Trim(lrsTab_Ped_Dizer_Pad("Dsc_Dizer_Endo"))
            Case 20
                textoEspec20 = Trim(lrsTab_Ped_Dizer_Pad("Dsc_Dizer_Endo"))
            End Select

            lrsTab_Ped_Dizer_Pad.MoveNext
        Loop
    End If

    'Ficha 1000854
    If gM46V999.gfAbrBasDados("P00MPED", gM46V999.gTipP00MPED, gM46V999.gEndP00MPED, lbdTab_Ped_Espec, _
                              pMensagem) = False Then Exit Function

    lNum_Pag = 0

    lSelect = "Select max(Num_Pag) as Nro_Pag from Tab_Ped_Espec where NOSSO_NUMERO = '" & pobjstcPed.NOSSO_NUMERO & "'"
    If gM46V999.gfObtRegistro(lbdTab_Ped_Espec, lSelect, lrsTab_Ped_Espec, pMensagem) Then
        If Not lrsTab_Ped_Espec.EOF Then
            lNum_Pag = IIf(IsNull(lrsTab_Ped_Espec("Nro_Pag")), 0, lrsTab_Ped_Espec("Nro_Pag"))
        End If
    End If

    Set ObjTextoEspec = New clsA46V726A
    Set ObjTextoEspec.M46V999 = gM46V999

    Set lobjstcTextoEspec = New stcA46V726B
    lobjstcTextoEspec.NOSSO_NUMERO = pobjstcPed.NOSSO_NUMERO
    lobjstcTextoEspec.Cod_Espec = lCodDizer
    lobjstcTextoEspec.Cod_Org = pobjstcPed.Cod_Org
    lobjstcTextoEspec.Cod_Ramo = pobjstcPed.Cod_Ramo
    lobjstcTextoEspec.Num_Ped = pobjstcPed.Num_Ped
    lobjstcTextoEspec.Num_Apol = pobjstcPed.Num_Apol
    lobjstcTextoEspec.Num_Endo = pobjstcPed.Num_Endo
    lobjstcTextoEspec.Num_pag = lNum_Pag + 1
    lobjstcTextoEspec.Qtd_vias = 1
    lobjstcTextoEspec.Espec_lin1 = textoEspec1
    lobjstcTextoEspec.Espec_lin2 = textoEspec2
    lobjstcTextoEspec.Espec_lin3 = textoEspec3
    lobjstcTextoEspec.Espec_lin4 = textoEspec4
    lobjstcTextoEspec.Espec_lin5 = textoEspec5
    lobjstcTextoEspec.Espec_lin6 = textoEspec6
    lobjstcTextoEspec.Espec_lin7 = textoEspec7
    lobjstcTextoEspec.Espec_lin8 = textoEspec8
    lobjstcTextoEspec.Espec_lin9 = textoEspec9
    lobjstcTextoEspec.Espec_lin10 = textoEspec10
    lobjstcTextoEspec.Espec_lin11 = textoEspec11
    lobjstcTextoEspec.Espec_lin12 = textoEspec12
    lobjstcTextoEspec.Espec_lin13 = textoEspec13
    lobjstcTextoEspec.Espec_lin14 = textoEspec14
    lobjstcTextoEspec.Espec_lin15 = textoEspec15
    lobjstcTextoEspec.Espec_lin16 = textoEspec16
    lobjstcTextoEspec.Espec_lin17 = textoEspec17
    lobjstcTextoEspec.Espec_lin18 = textoEspec18
    lobjstcTextoEspec.Espec_lin19 = textoEspec19
    lobjstcTextoEspec.Espec_lin20 = textoEspec20
    lobjstcTextoEspec.NUM_PI = pobjstcPed.NUM_PI

    If Not pobjstcPed.FindTextoEspec(lobjstcTextoEspec) Then   'Ficha 1000837 - Triagem Informar Beneficiario nos Itens
        If Not ObjTextoEspec.mfIncluir(lobjstcTextoEspec, "", "", "") Then
            pMensagem = ObjTextoEspec.UltimoErro
        Else   'Ficha 1000837 - Triagem Informar Beneficiario nos Itens
            pobjstcPed.ColecaoTextoEspec.Add lobjstcTextoEspec, lobjstcTextoEspec.Class_ID   'Ficha 1000837 - Triagem Informar Beneficiario nos Itens
        End If
    End If   'Ficha 1000837 - Triagem Informar Beneficiario nos Itens

    Set ObjTextoEspec = Nothing

    mfProcuraDizer = True

Finalizar:

    Exit Function

mfProcuraDizer_Err:
    MsgBox "Erro na fun��o mfProcuraDizer - Contactar Sistemas"
End Function

Function mfDeletaDizer(ByRef pobjstcPed As stcA46V702B, lCodDizer As Integer, ByRef pMensagem As String) As Boolean
    Dim objPedido As clsA46V702A    'Objeto da classe clsA46V702A  (TAB_PED)
    Dim objTabVenc As clsA46V713A
    Dim lbdTab_Ped_Dizer_Pad As ADODB.Connection
    Dim lrsTab_Ped_Dizer_Pad As clsYasRecordSet
    Dim lbdBasDados As ADODB.Connection    'Base de dados.
    Dim lrsTabela As clsYasRecordSet    'Registro da tabela.
    Dim lobjstcTextoEspec As stcA46V726B   'Objeto da classe de estrutura stcA46V726B (TAB_PED_ESPEC).
    Dim ObjTextoEspec As clsA46V726A    'Objeto da classe de estrutura clsA46V726A (TAB_PED_ESPEC).
    Dim lObjstcVenc As stcA46V713B    'Objeto da classe de estrutura stcA46V713B (TAB_PED_Venc)
    Dim lSql As String
    Dim lI As Integer
    Dim lNum_pi As Long
    Dim lNosNumero As String

    On Error GoTo mfDeletaDizer_Err

    mfDeletaDizer = False

    Set ObjTextoEspec = New clsA46V726A
    Set ObjTextoEspec.M46V999 = gM46V999

    Set lobjstcTextoEspec = New stcA46V726B
    lobjstcTextoEspec.NOSSO_NUMERO = pobjstcPed.NOSSO_NUMERO
    lobjstcTextoEspec.Cod_Espec = lCodDizer
    If Not ObjTextoEspec.mfExcluir(lobjstcTextoEspec, "", "", "") Then
        pMensagem = ObjTextoEspec.UltimoErro
    End If
    Set ObjTextoEspec = Nothing

    mfDeletaDizer = True

Finalizar:

    Exit Function

mfDeletaDizer_Err:
    MsgBox "Erro na fun��o mfDeletaDizer - Contactar Sistemas"
End Function

Private Function ReentradaImobiliario(ByVal pNum_pi As Long, ByVal pNosso_Numero As String, ByRef pNosso_Numero_Novo As String, ByVal pMensagem As String) As Boolean
    Dim lBanco As ADODB.Connection
    Dim lrs As clsYasRecordSet
    Dim lSelect As String
    
    ReentradaImobiliario = False
    pNosso_Numero_Novo = Empty
    
    If gM46V999.gfAbrBasDados("RAMOSDIVERSOS", gM46V999.gTipDIVERSOS, gM46V999.gEndDIVERSOS, lBanco, pMensagem) = False Then
        MsgBox pMensagem, vbCritical
        GoTo SAIDA
    End If
    
    lSelect = "SELECT Num_PI FROM Ftp_Tab_ped WHERE COD_ESTIP = 20 AND Nosso_Numero = " & pNosso_Numero
    
    If Not gM46V999.gfObtRegistro(lBanco, lSelect, lrs, pMensagem) Then
        GoTo SAIDA
    End If
    
    If lrs.EOF Then
        pMensagem = "Este PI n�o existe ou n�o � Imobiliario!"
        GoTo SAIDA
    End If
    
    'Limpa Transmiss�es anteriores
    lSelect = "DELETE From tab_ctrl_transmissao WHERE num_pi = " & pNum_pi
    
    If gM46V999.gfExeSQL(lBanco, lSelect, pMensagem) > 0 Then
        Call gM46V999.gpGraLog(1, pMensagem)
        GoTo SAIDA
    End If
    
    lSelect = "EXEC proc_StartFluxoEmissaoTransmissaoMulti_Kit '" & pNosso_Numero & "'," & pNum_pi
    
    If gM46V999.gfExeSQL(lBanco, lSelect, pMensagem) > 0 Then
        Call gM46V999.gpGraLog(1, pMensagem)
        GoTo SAIDA
    End If
    
    'Impede Transmiss�o automatica
    lSelect = "UPDATE Tab_Ped SET COD_POSI_ATUAL = 0 WHERE num_pi = " & pNum_pi
    
    If gM46V999.gfExeSQL(lBanco, lSelect, pMensagem) > 0 Then
        Call gM46V999.gpGraLog(1, pMensagem)
        GoTo SAIDA
    End If
    
    lSelect = "SELECT Nosso_Numero FROM Tab_Ped WHERE Num_Pi = " & pNum_pi
    If gM46V999.gfObtRegistro(lBanco, lSelect, lrs, pMensagem) = False Then
        Call gM46V999.gpGraLog(1, pMensagem)
        GoTo SAIDA
    End If
    
    If Not lrs.EOF Then
        pNosso_Numero_Novo = "" & lrs!NOSSO_NUMERO
        ReentradaImobiliario = True
    Else
        Call gM46V999.gpGraLog(1, "Erro ao executar a proc_StartFluxoEmissaoTransmissaoMulti_Kit!")
        GoTo SAIDA
    End If

SAIDA:
    GoSub FechaObj
    Exit Function
FechaObj:
    Set lrs = Nothing
    Return
End Function


'Private Sub updAtualizaStatusInspecao(Num_PI As String, ByRef pMensInsp As String)
'Dim lretorno    As String
'Dim lI          As Integer
'Dim lCod        As Integer
'Dim lAuxiliar   As String
'Dim lBDP00Multi     As ADODB.Connection
'Dim lRstErro        As clsYasRecordSet
'Dim lMensagem       As String
'Dim lSelect         As String
'Dim lrsPadrao       As clsYasRecordSet
'Dim NossoNumero     As String
'
'    pMensInsp = ""
'
'    If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lBDP00Multi, lMensagem) = False Then
'        Call gM46V999.gpGraLog(2, lMensagem)
'        Exit Sub
'    End If
'
'    ' Tentar encontrar a Inspe��o a partir do PI e Triagem
'    lSelect = "SELECT 1 " & vbCrLf
'    lSelect = lSelect & "FROM        Tab_Trans_Insp   TTI " & vbCrLf
'    lSelect = lSelect & "INNER JOIN  ged..Tab_Triagem                TT " & vbCrLf
'    lSelect = lSelect & "    ON      TTI.NOSSO_NUMERO                = TT.num_siscota " & vbCrLf
'    lSelect = lSelect & "        AND TT.num_pi                       = '" & Num_PI & "' AND TTI.Cod_Status_I = '0' " & vbCrLf
'
'    If gM46V999.gfObtRegistro(lBDP00Multi, lSelect, lrsPadrao, lMensagem) Then
'        If Not lrsPadrao.EOF Then
'    lSelect = "UPDATE      TTI " & vbCrLf
'    lSelect = lSelect & "SET         Cod_Status_I                    = '1' " & vbCrLf
'    lSelect = lSelect & "FROM        Tab_Trans_Insp   TTI " & vbCrLf
'    lSelect = lSelect & "INNER JOIN  ged..Tab_Triagem                TT " & vbCrLf
'    lSelect = lSelect & "    ON      TTI.NOSSO_NUMERO                = TT.num_siscota " & vbCrLf
'    lSelect = lSelect & "        AND TT.num_pi                       = '" & Num_PI & "' AND TTI.Cod_Status_I = '0' " & vbCrLf
'
'    If gM46V999.gfExeSQL(lBDP00Multi, lSelect, lMensagem) <> 0 Then
'        Call gM46V999.gpGraLog(2, lMensagem)
'        Exit Sub
'    End If
'        Else
'            ' N�o encontrou, tentar ent�o obter o nosso n�mero a partir da TAB_PED
'            lSelect = "SELECT      TPNC.NOSSO_NUMERO_OLD " & vbCrLf
'            lSelect = lSelect & "FROM        Tab_Ped                 TP " & vbCrLf
'            lSelect = lSelect & "INNER JOIN  Tab_Ped_Num_Copia       TPNC " & vbCrLf
'            lSelect = lSelect & "    ON      TPNC.NOSSO_NUMERO_NEW   = TP.NOSSO_NUMERO " & vbCrLf
'            lSelect = lSelect & "INNER JOIN  Tab_Trans_Insp          TTI " & vbCrLf
'            lSelect = lSelect & "    ON      TTI.NOSSO_NUMERO        = TPNC.NOSSO_NUMERO_OLD " & vbCrLf
'            lSelect = lSelect & "WHERE       TP.NUM_PI       = '" & Num_PI & "' "
'
'            If gM46V999.gfObtRegistro(lBDP00Multi, lSelect, lrsPadrao, lMensagem) Then
'                If Not lrsPadrao.EOF Then
'                    NossoNumero = lrsPadrao("NOSSO_NUMERO_OLD")
'                    lSelect = "UPDATE Tab_Trans_Insp " & vbCrLf
'                    lSelect = lSelect & "SET Cod_Status_I = '1' " & vbCrLf
'                    lSelect = lSelect & "WHERE NOSSO_NUMERO = '" & NossoNumero & "' AND Cod_Status_I = '0'  " & vbCrLf
'
'                    If gM46V999.gfExeSQL(lBDP00Multi, lSelect, lMensagem) <> 0 Then
'                        Call gM46V999.gpGraLog(2, lMensagem)
'                        Exit Sub
'                    End If
'
'                End If
'            End If
'        End If
'    End If
'
'
'    Call gM46V999.gpFecha2(lBDP00Multi)
'
'End Sub
