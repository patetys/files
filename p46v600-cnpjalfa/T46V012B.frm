VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmT46V012B 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dados de transmiss�o"
   ClientHeight    =   5745
   ClientLeft      =   420
   ClientTop       =   1050
   ClientWidth     =   10140
   Icon            =   "T46V012B.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   10140
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel panAviso 
      Height          =   5745
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10155
      _Version        =   65536
      _ExtentX        =   17912
      _ExtentY        =   10134
      _StockProps     =   15
      ForeColor       =   12582912
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      BevelInner      =   1
      Outline         =   -1  'True
      Begin VB.TextBox txtPagAntBloqueto 
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   5760
         MaxLength       =   9
         TabIndex        =   9
         ToolTipText     =   "N�mero de bloqueto do pagamento antecipado."
         Top             =   5040
         Width           =   1000
      End
      Begin VB.PictureBox picOk 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         Height          =   255
         Left            =   1200
         Picture         =   "T46V012B.frx":030A
         ScaleHeight     =   195
         ScaleWidth      =   225
         TabIndex        =   4
         Top             =   2160
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.PictureBox picChecked 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         Height          =   255
         Left            =   840
         Picture         =   "T46V012B.frx":0404
         ScaleHeight     =   195
         ScaleWidth      =   225
         TabIndex        =   3
         Top             =   2160
         Visible         =   0   'False
         Width           =   285
      End
      Begin Threed.SSPanel panSiscad 
         Height          =   615
         Left            =   240
         TabIndex        =   1
         Top             =   4920
         Width           =   1935
         _Version        =   65536
         _ExtentX        =   3413
         _ExtentY        =   1085
         _StockProps     =   15
         Caption         =   "<-- Siscad      "
         ForeColor       =   12582912
         BackColor       =   13160660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   4
         Begin Threed.SSCommand cmdSiscad 
            Height          =   615
            Left            =   0
            TabIndex        =   2
            Top             =   0
            Width           =   735
            _Version        =   65536
            _ExtentX        =   1296
            _ExtentY        =   1085
            _StockProps     =   78
            Picture         =   "T46V012B.frx":04FE
         End
      End
      Begin MSFlexGridLib.MSFlexGrid flxResultado 
         Height          =   4620
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   8149
         _Version        =   393216
         Rows            =   1
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   -2147483624
         ForeColor       =   10485760
         BackColorBkg    =   -2147483624
         GridLines       =   0
         GridLinesFixed  =   0
         AllowUserResizing=   1
         BorderStyle     =   0
         Appearance      =   0
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
      Begin Threed.SSCommand cmdVoltar 
         Height          =   300
         Left            =   9000
         TabIndex        =   6
         Top             =   5160
         Width           =   1005
         _Version        =   65536
         _ExtentX        =   1773
         _ExtentY        =   529
         _StockProps     =   78
         Caption         =   "&Voltar"
         ForeColor       =   12582912
      End
      Begin Threed.SSCommand cmdEmitir 
         Height          =   300
         Left            =   7920
         TabIndex        =   7
         Top             =   5160
         Width           =   1005
         _Version        =   65536
         _ExtentX        =   1773
         _ExtentY        =   529
         _StockProps     =   78
         Caption         =   "&Emitir"
         ForeColor       =   12582912
      End
      Begin VB.Label lblboleto 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bloqueto da transmiss�o:"
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   3720
         TabIndex        =   10
         ToolTipText     =   "N�mero de bloqueto do pagamento antecipado."
         Top             =   5040
         Width           =   1965
      End
      Begin VB.Label lblAviso 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   3000
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmT46V012B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mComUltData As Double
Private mComUltUsuario As String
Private mRetorno As Integer
Private mobjstcPed As stcA46V702B
Private mobjstcPedLoc As stcA46V708B
Public mErroSiscad As Boolean
Public mErroDemais As Boolean
Private mSuspProduto As Boolean
Private bdDiversos As ADODB.Connection

Public Sub mpInstanciarForm(ByRef pobjstcPed As stcA46V702B, pSiscad As Boolean, _
                            pDemais As Boolean, ByRef pRetorno As Integer)
    Dim lMensagem As String
    Set mobjstcPed = pobjstcPed
    If pSiscad = True Then
        panSiscad.Visible = True
    Else
        panSiscad.Visible = False
    End If
    mErroSiscad = pSiscad
    mErroDemais = pDemais

    If pobjstcPed.Tip_Cust_Emis = 5 Or pobjstcPed.Tip_Cust_Emis = 4 Then
        lblboleto.Visible = True
        txtPagAntBloqueto.Visible = True
        txtPagAntBloqueto.Text = pobjstcPed.objstcCtrlEmis.Num_Bloq & "-" & pobjstcPed.objstcCtrlEmis.DIG_BLOQUETO
    Else
        lblboleto.Visible = False
        txtPagAntBloqueto.Visible = False
        txtPagAntBloqueto.Text = ""
    End If

    If gM46V999.gfAbrBasDados("RAMOSDIVERSOS", gM46V999.gTipDIVERSOS, gM46V999.gEndDIVERSOS, bdDiversos, lMensagem) = False Then
        Exit Sub
    End If
    MousePointer = vbDefault
    Me.Show vbModal
    pRetorno = mRetorno
End Sub
Private Sub cmdEmitir_Click()
    Dim lTipoTransmissao As Long
    Dim bdBasDados As ADODB.Connection
    Dim lobjCtrlEmis As clsA46V712A    'Objeto da classe clsA46V712A (TAB_CTRL_EMIS)
    Dim lObjtabPed As clsA46V702A
    Dim lobjPedLoc As clsA46V708A
    Dim lResultado As Integer
    Dim lOK As String
    Dim lAcao As Integer             'Fluxo
    Dim lAuxiliar As String
    Dim clsA99V001 As clsA99V001
    Dim lDigito As String
    Dim lMensagem As String

    mobjstcPed.NUM_PI = Format(frmT46V012A.txtNumeroPI.Text, "0000000000")
    mobjstcPed.objstcCtrlEmis.Dat_Rec_Cia = gM46V999.gfForData(1, frmT46V012A.txtDataRecebimento.Text)
    mobjstcPed.objstcCtrlEmis.NUM_PI = mobjstcPed.NUM_PI
    mobjstcPed.Cod_Unicidade = val(frmT46V012A.txtDV.Text)
    mobjstcPed.Dat_Pri_Calculo = 0
    mobjstcPed.Cod_Tip_Doc = e_TipDoc_Proposta
    If mobjstcPed.Cod_Bco_Cobr = 347 Then
        mobjstcPed.Cod_Bco_Cobr = 356
    End If

    If txtPagAntBloqueto.Visible = True Then
        lAuxiliar = Replace(txtPagAntBloqueto.Text, "-", "")
        Set clsA99V001 = New clsA99V001
        Set clsA99V001.M46V999 = gM46V999

        Call clsA99V001.DIG_BLOQUETO(Format(lAuxiliar, "000000000"), lDigito, lMensagem)
        
        ' Valida��o removida devido ao projeto de boleto registrado
'        If Right(lAuxiliar, 1) <> lDigito Then
'            Call gM46V999.gpGraLog(1, "FGE0059 - Pagamento antecipado com bloqueto - n�mero do bloqueto com d�gito inv�lido.")
'            Exit Sub
'        End If
        mobjstcPed.objstcCtrlEmis.Num_Bloq = Left(Format(val(Replace(txtPagAntBloqueto.Text, "-", "")), "00000000"), "7")
        mobjstcPed.objstcCtrlEmis.DIG_BLOQUETO = val(Right(txtPagAntBloqueto.Text, 1))
        Set clsA99V001 = Nothing
    End If


    gM46V111.gNum_Pi = mobjstcPed.NUM_PI

Retorno:
    If gM46V999.gEfetuouCalculo = True Then
        'Verifico se existem erros de triagem que foram detectados,
        ' se tiver ent�o eu n�o permito emitir, somente suspender
        If mfExisteErrosTriagem Then
            gM46V999.gEfetuouCalculo = False
        End If
    End If

    Call frmT46V154A.mpInstanciarForm(lTipoTransmissao, lAcao)

    If lTipoTransmissao > 0 Then
        Screen.MousePointer = vbHourglass
        'Call gpTransmitirPropostaOCX(lTipoTransmissao, lResultado, objPedido)
        Select Case lTipoTransmissao
        Case 1
            mobjstcPed.Cod_Posi_Atual = e_IdeEmitida_Emitir
            lOK = "transmitido"
        Case 2
            mobjstcPed.Cod_Posi_Atual = e_IdeEmitida_Suspender
            lOK = "suspenso"
        Case 3
            mobjstcPed.Cod_Posi_Atual = e_IdeEmitida_EmitirProd
            lOK = "emitido produto"
        Case 5
            mobjstcPed.Cod_Posi_Atual = e_IdeEmitida_SuspenderProd
            lOK = "suspenso produto"
        Case 6
            mobjstcPed.Cod_Posi_Atual = e_IdeEmitida_Reemissao
            lOK = "reemiss�o"
        Case Else
            mobjstcPed.Cod_Posi_Atual = e_IdeEmitida_Nao
            lOK = "recusado"
        End Select

        mobjstcPed.DAT_DIGIT = Left(gM46V999.gfForData(3), 8)

        Set lObjtabPed = New clsA46V702A
        Set lObjtabPed.ConexaoBD = bdBasDados
        Set lObjtabPed.M46V999 = gM46V999

        If Not lObjtabPed.mfAlterar(mobjstcPed, "", "", mobjstcPed.Dat_Liber, mobjstcPed.Cod_User_Liber) Then
            Call gM46V999.gpGraLog(2, lObjtabPed.UltimoErro)
            Exit Sub
        End If

        mobjstcPed.objstcCtrlEmis.Dat_Digit_MF = Left(gM46V999.gfForData(3), 8)
        mobjstcPed.objstcCtrlEmis.Cod_Iden_Stat_Acao = 0

        Set lobjCtrlEmis = New clsA46V712A
        Set lobjCtrlEmis.ConexaoBD = bdBasDados
        Set lobjCtrlEmis.M46V999 = gM46V999

        If Not lobjCtrlEmis.mfAlterar(mobjstcPed.objstcCtrlEmis, "", "") Then
            Call gM46V999.gpGraLog(2, lobjCtrlEmis.UltimoErro)
            Exit Sub
        End If

        Screen.MousePointer = vbDefault

        'Se ocorreu tudo OK finalizar o sistema
        If lOK <> "" Then
            MsgBox "Pi: " & mobjstcPed.objstcCtrlEmis.NUM_PI & " ser� " & lOK & ".", vbInformation
            mRetorno = lTipoTransmissao
            Unload Me

            Exit Sub
        End If
    End If

End Sub

Private Sub cmdSiscad_Click()
    Call panSiscad_Click
End Sub

Private Sub cmdVoltar_Click()
    Unload Me
    mRetorno = 0
End Sub


Private Sub panSiscad_Click()

    If UCase(gM46V999.gSerDIVERSOS) Like "*MIX*" Then
        gM46V999.gEndExeSiscad = Replace(gM46V999.gEndExeSiscad, "P37V000", "P37V001")
    End If

    Dim lCgcCpfSiscad As String    'CGC/CPF formatado p/ par�metro do Siscad (suporta CNPJ alfanum�rico).

    If IsNumeric(mobjstcPed.Num_Cgc_Cpf) Then
        lCgcCpfSiscad = Format(mobjstcPed.Num_Cgc_Cpf, String(14, "0"))
    Else
        lCgcCpfSiscad = Right$(String(14, "0") & Trim(mobjstcPed.Num_Cgc_Cpf), 14)
    End If

    Call Shell(gM46V999.gEndExeSiscad & " " & lCgcCpfSiscad & "0" & _
               Format(frmT46V012A.txtNumeroPI.Text, String(10, "0")), vbNormalFocus)

    'Ficha 901363
    'If MsgBox("Deseja atualizar os dados da proposta com os dados do siscad??", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then

    Call mpAtualizaDadosSegurado

    'End If
End Sub

Public Sub mpAtualizaDadosSegurado()
    Dim lCPF_CNPJ As String
    Dim rsDiversos As clsYasRecordSet
    Dim rsSyasCob As clsYasRecordSet
    Dim lMensagem As String
    Dim lLibAltEnd As Boolean
    Dim lLibAltNome As Boolean
    Dim lObjPed As clsA46V702A
    Dim lobjPedEnder As clsA46V707A
    Dim lSelect As String

    lSelect = " select * from p0044800.dbo.t0044800 Where "
    lSelect = lSelect & " num_cgc_cpf = '" & Trim(mobjstcPed.Num_Cgc_Cpf) & "'"

    If gM46V999.gfObtRegistro(bdDiversos, lSelect, rsDiversos, lMensagem) = False Then
        Call gM46V999.gpFecha2(bdDiversos)
        Call gM46V999.gpGraLog(1, lMensagem)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    If Not rsDiversos.EOF Then
        mobjstcPed.objstcPedEndereco.NOSSO_NUMERO = mobjstcPed.NOSSO_NUMERO
        mobjstcPed.Nom_Segurado = Trim(Left(Trim(rsDiversos("NOME_RAZAO")), 60))

        If gM46V999.gfPreenchido(Trim(rsDiversos("NOM_LOCTRAB"))) Then
            mobjstcPed.Nom_Segurado = Left(mobjstcPed.Nom_Segurado & " (" & Trim(rsDiversos("NOM_LOCTRAB")) & ") ", 60)
        End If

        mobjstcPed.objstcPedEndereco.Cod_Tip_Logr = Trim(rsDiversos("COD_TIP_LOGR"))
        mobjstcPed.objstcPedEndereco.Nom_Logr = Trim(rsDiversos("NOM_LOGR"))
        mobjstcPed.objstcPedEndereco.Dsc_Compl_Logr = Trim(rsDiversos("DSC_COMPL_LOGR"))
        mobjstcPed.objstcPedEndereco.Nom_Bairro = Trim(rsDiversos("NOM_BAIRRO"))
        mobjstcPed.objstcPedEndereco.Nom_Cid = Trim(rsDiversos("NOM_CID"))

        If Trim(rsDiversos("SIG_UF")) <> "" Then
            mobjstcPed.objstcPedEndereco.Sig_UF = Trim(rsDiversos("SIG_UF"))
        End If

        mobjstcPed.objstcPedEndereco.Num_CEP = Format(rsDiversos("Num_CEP"), "00000000")
        '1001046 - Telefone do Segurado 8 Digitos
        mobjstcPed.objstcPedEndereco.Num_Fone = Mid(rsDiversos("NUM_FONE"), 1, 8)

        'txtCodSegurado.Text = rsDiversos("COD_IDENTIFICADOR")
        ''''TIP_PESSOA - jur = 1 fis = 2

        mobjstcPed.Tip_Pessoa = Trim(rsDiversos("TIP_PESSOA"))
        mobjstcPed.Cod_Ramo_Ativ = mobjstcPed.gCodigoRamoAtiv(rsDiversos("RAMO_ATIV_EMPR"))

        '[FRH] - 14/05/2009 - DE COD_TIP_DOC P/ COD_DOC
        mobjstcPed.Cod_Doc = Trim(rsDiversos("TIP_DOCUM"))

        mobjstcPed.Num_Doc = Trim(rsDiversos("NUM_DOCUM"))
        mobjstcPed.Org_Exp = Trim(rsDiversos("NOM_ORG_EXPED"))
        mobjstcPed.Dat_Emis_Doc = Trim(rsDiversos("DAT_EXPED"))

        'Ficha 900417
        If gM46V999.gCodProduto >= 968 Then
            mobjstcPed.Nom_Empresa = Trim(rsDiversos("NOM_PROFISSAO"))
        Else
            mobjstcPed.Nom_Empresa = Trim(rsDiversos("NOM_LOCTRAB"))
        End If

        mobjstcPed.Sexo = Trim(rsDiversos("SEXO"))
        mobjstcPed.Val_RendaFatu = gM46V999.gfTratarNulo(rsDiversos("RENDAFATU"), TpoValor)

        lSelect = " select * from p0044800.dbo.T0044807 Where "
        lSelect = lSelect & " num_pi = " & frmT46V012A.txtNumeroPI.Text
        lSelect = lSelect & " and num_cgc_cpf = '" & Trim(mobjstcPed.Num_Cgc_Cpf) & "'"

        If gM46V999.gfObtRegistro(bdDiversos, lSelect, rsSyasCob, lMensagem) = False Then
            Call gM46V999.gpFecha2(bdDiversos)
            Call gM46V999.gpGraLog(1, lMensagem)
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        If Not rsSyasCob.EOF Then
            mobjstcPed.Cod_Tip_Logr = Trim(rsSyasCob("COD_TIP_LOGR"))
            mobjstcPed.Nom_Logr = Trim(rsSyasCob("NOM_LOGR"))
            mobjstcPed.Dsc_Compl_Logr = Trim(rsSyasCob("DSC_COMPL_LOGR"))
            mobjstcPed.Nom_Bairro = Trim(rsSyasCob("NOM_BAIRRO"))
            mobjstcPed.Nom_Cid = Trim(rsSyasCob("NOM_CID"))
            mobjstcPed.Num_CEP = Format(rsSyasCob("Num_CEP"), "00000000")

            If Trim(rsSyasCob("SIG_UF")) <> "" Then
                mobjstcPed.Sig_UF = Trim(rsSyasCob("SIG_UF"))
            End If
            If Not IsNull(rsSyasCob("num_fone")) Then
                '1001046 - Telefone do Segurado 8 Digitos
                mobjstcPed.Num_Fone = Mid(rsSyasCob("num_fone"), 1, 8)
            Else
                mobjstcPed.Num_Fone = 0
            End If
        End If

        Set lObjPed = New clsA46V702A
        Set lObjPed.M46V999 = gM46V999

        If Not lObjPed.mfAlterar(mobjstcPed, "", "") Then
            Call gM46V999.gpGraLog(1, lObjPed.UltimoErro)
            Exit Sub
        End If

        Set lobjPedEnder = New clsA46V707A
        Set lobjPedEnder.M46V999 = gM46V999

        If Not lobjPedEnder.mfAlterar(mobjstcPed.objstcPedEndereco, "", "") Then
            Call gM46V999.gpGraLog(1, lobjPedEnder.UltimoErro)
            Exit Sub
        End If

        Call gM46V999.gpGraLog(1, "Dados atualizados com sucesso!")
        If mErroDemais = False Then
            gM46V999.gEfetuouCalculo = True
            Call cmdEmitir_Click
        End If

    End If
End Sub
Private Function mfExisteErrosTriagem() As Boolean
    Dim lSelect As String       'Select
    Dim lrst As clsYasRecordSet    'Registro
    Dim lMensagem As String         'Mensagem

    lSelect = "Select * from ged.dbo.tab_prop_mensagem where num_pi = " & gM46V111.gNum_Pi
    lSelect = lSelect & " and dat_liber=0"

    If gM46V999.gfObtRegistro(bdDiversos, lSelect, lrst, lMensagem) = False Then
        Call gM46V999.gpGraLog(1, lMensagem)
        Exit Function
    End If
    If lrst.EOF Then
        mfExisteErrosTriagem = False
    Else
        mfExisteErrosTriagem = True
    End If
End Function

Private Sub txtPagAntBloqueto_GotFocus()
    With txtPagAntBloqueto
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub
Private Sub txtPagAntBloqueto_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub txtPagAntBloqueto_LostFocus()
    Dim lAuxiliar As String     'Auxiliar.

    With txtPagAntBloqueto
        If gM46V999.gfPreenchido(.Text) Then
            .Text = Trim$(.Text)
            If Len(.Text) = 9 Then
                Exit Sub
            End If
            If IsNumeric(.Text) Then
                lAuxiliar = Format$(CDbl(.Text), "00000000")
                .Text = Mid$(lAuxiliar, 1, 7) & "-" & Mid$(lAuxiliar, 8, 1)
            End If
        End If
    End With
End Sub
