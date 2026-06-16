VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form frmT46V145A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Itens para Endosso"
   ClientHeight    =   5805
   ClientLeft      =   4035
   ClientTop       =   2355
   ClientWidth     =   6840
   Icon            =   "T46V145A.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fmTipoEndosso 
      Caption         =   "Tipo de endosso:"
      ForeColor       =   &H00800000&
      Height          =   810
      Left            =   30
      TabIndex        =   7
      Top             =   4470
      Width           =   6735
      Begin Threed.SSOption optTipoEndosso 
         Height          =   315
         Index           =   0
         Left            =   1515
         TabIndex        =   8
         Top             =   300
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   556
         _StockProps     =   78
         Caption         =   "Altera��es"
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.26
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optTipoEndosso 
         Height          =   315
         Index           =   2
         Left            =   3810
         TabIndex        =   9
         Top             =   300
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   556
         _StockProps     =   78
         Caption         =   "Exclus�o de item"
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.26
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin Threed.SSCommand cmdOk 
      Default         =   -1  'True
      Height          =   300
      Left            =   2385
      TabIndex        =   4
      ToolTipText     =   "Volta � janela de item."
      Top             =   5430
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Ok"
      ForeColor       =   12582912
   End
   Begin Threed.SSCommand cmdCancelar 
      Cancel          =   -1  'True
      Height          =   300
      Left            =   3465
      TabIndex        =   5
      ToolTipText     =   "Volta � janela de item."
      Top             =   5430
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Cancelar"
      ForeColor       =   12582912
   End
   Begin VB.Frame fraOpcao 
      Height          =   2895
      Left            =   30
      TabIndex        =   0
      Top             =   45
      Width           =   6735
      Begin VB.OptionButton optItens 
         Caption         =   "Listar todos os itens"
         ForeColor       =   &H00C00000&
         Height          =   495
         Index           =   0
         Left            =   1320
         TabIndex        =   3
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton optItens 
         Caption         =   "Item:"
         ForeColor       =   &H00C00000&
         Height          =   495
         Index           =   1
         Left            =   2730
         TabIndex        =   2
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtNumItem 
         Height          =   375
         Left            =   3450
         MaxLength       =   4
         TabIndex        =   1
         Top             =   1500
         Width           =   855
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flxConsulta 
      Height          =   4335
      Left            =   30
      TabIndex        =   6
      Top             =   60
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   7646
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      BackColor       =   16777215
      ForeColor       =   12582912
      BackColorFixed  =   12632256
      ForeColorFixed  =   12582912
      BackColorSel    =   8388608
      ForeColorSel    =   16777215
      BackColorBkg    =   12632256
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.Label lblAviso 
      Caption         =   "*Para endosso de inclus�o, informe o n�mero 9999."
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   4425
      TabIndex        =   10
      Top             =   5310
      Width           =   2055
   End
End
Attribute VB_Name = "frmT46V145A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mNumeroApolice As Double                    'N�mero da Ap�lice
Dim mNumerosItens As String                 'N�meros dos itens desta ap�lice pertencentes a esta cota��o

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim lobjstcPedLoc As stcA46V708B
    Dim lobjPedLoc As clsA46V708A
    Dim lNumItem As String
    Dim lMensagem As String
    Dim lCodDizer As Byte

    lNumItem = flxConsulta.TextMatrix(flxConsulta.Row, 0)

    If Not IsNumeric(lNumItem) Then Exit Sub

    lCodDizer = 0


    If optTipoEndosso(0).Value = True Then
        lCodDizer = 12
    End If

    If optTipoEndosso(2).Value = True Then
        lCodDizer = 4
    End If
    If lCodDizer = 0 Then
        Call gM46V999.gpGraLog(1, "Informe o tipo de endosso!")
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    If gM46V001.gfGraEndossoYasuda(frmT46V101A.objstcPedLoc.Num_Apol_Anterior, lNumItem, frmT46V101A.objstcPedido.NOSSO_NUMERO, _
                                   e_SimNao_Nao, lCodDizer, lMensagem, False, frmT46V101A.objstcPedido, "") = False Then
        If gM46V999.gfPreenchido(lMensagem) Then
            Call gM46V999.gpGraLog(1, lMensagem)
        End If
        Exit Sub
    End If

    Set lobjstcPedLoc = New stcA46V708B
    lobjstcPedLoc.NOSSO_NUMERO = frmT46V101A.objstcPedido.NOSSO_NUMERO
    lobjstcPedLoc.Num_Item = lNumItem
    Set lobjPedLoc = New clsA46V708A
    Set lobjPedLoc.M46V999 = gM46V999
    If Not lobjPedLoc.mfCarregarColecao(lobjstcPedLoc, "", "", True) Then
        Call gM46V999.gpGraLog(1, lobjPedLoc.UltimoErro)
        Exit Sub
    End If
    If lobjPedLoc.Colecao.Count <> 1 Then
        Call gM46V999.gpGraLog(1, "erro ao criar o endosso")
        Exit Sub
    End If
    Set lobjstcPedLoc = lobjPedLoc.Colecao.Item(1)
    lobjstcPedLoc.Cod_Ramo = gM46V999.gCod_Ramo
    lobjstcPedLoc.Cod_Iden_Tarif = gM46V999.gCodProduto
    lobjstcPedLoc.Cod_DizerItem = lCodDizer
    lobjstcPedLoc.Num_Item_Ant = lNumItem
    Set lobjstcPedLoc.ObjPai = frmT46V101A.objstcPedido
    frmT46V101A.objstcPedido.ColecaoPedLoc.Add lobjstcPedLoc, lobjstcPedLoc.Class_ID

    Set frmT46V101A.objstcPedLoc = frmT46V101A.objstcPedido.ColecaoPedLoc.Item(lobjstcPedLoc.Class_ID)
    
    Call frmT46V101A.objstcPedido.addItemPedidoApolice(lobjstcPedLoc)
    
    gCarga = True
    'Posiciona o item.
    Call frmT46V101A.mpPosItem(lobjstcPedLoc)

    '6. Atualiza grade de itens.
    'Call mpCriGraItens

    'Call BuscaItemEndosso(frmT46V101A.objstcPedLoc.Num_Apol_Anterior, lNumItem)

    Screen.MousePointer = vbDefault

    Me.Hide
End Sub

Public Sub BuscaItemEndosso(txtRenApolice As String, txtRenItem As String)
    Dim lbdP00APOL As ADODB.Connection    'Arquivo - P00APOL.
    Dim lBonus As Long
    Dim lCodSeguradora As Integer     'C�digo da seguradora.
    Dim lIndSinistro As Byte      'Indicador de sinistro.
    Dim lNumItem As Integer    'N�mero do item.
    Dim lMensagem As String    'Mensagem.
    Dim rsAPOL As clsYasRecordSet    'Registro: P00APOL.
    Dim lI As Long    'Auxiliar.
    Dim lobjstcClau As stcA46V703B    'Objeto da classe de estrutura stcA46V703B (TAB_PED_CLAU)
    Dim ColecaoClau As Collection    'Cole��o de cl�usulas.
    Dim lobjstcCobert As stcA46V704B    'Objeto da classe de estrutura stcA46V704B (TAB_PED_COBERT)
    Dim ColecaoCobert As Collection    'Cole��o de coberturas.
    Dim lRegistro As String    'Registro.
    Dim lIdeTabela As String
    Dim lSelect As String

    Screen.MousePointer = vbHourglass


    '2. Consiste n�mero de ap�lice anterior.
    If Not gM46V999.gfPreenchido(txtRenApolice) Then
        Call gM46V999.gpGraLog(1, "FGE0069 - " & gM46V999.gPrefixo & _
                                  "N�mero da ap�lice inv�lido (n�o preenchido).")
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If Not IsNumeric(txtRenApolice) Then
        Call gM46V999.gpGraLog(1, "FGE0069 - " & gM46V999.gPrefixo & _
                                  "N�mero da ap�lice inv�lido (n�o num�rico).")
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If Val(txtRenApolice) = 0 Then
        Call gM46V999.gpGraLog(1, "FGE0069 - " & gM46V999.gPrefixo & _
                                  "N�mero da ap�lice inv�lido (zerado).")
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If Len(Trim(txtRenApolice)) < 10 Then
        txtRenApolice = Trim$(txtRenApolice)
        If Len(txtRenApolice) < 10 Then
            txtRenApolice = String(10 - Len(txtRenApolice), "0") & txtRenApolice
        End If
    End If

    '3. Consiste n�mero do item da ap�lice.
    If Not gM46V999.gfPreenchido(txtRenItem) Then
        Call gM46V999.gpGraLog(1, "FGE0070 - " & gM46V999.gPrefixo & _
                                  "N�mero do item da ap�lice inv�lido ( n�o preenchido).")
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If Not IsNumeric(txtRenItem) Then
        Call gM46V999.gpGraLog(1, "FGE0070 - " & gM46V999.gPrefixo & _
                                  "N�mero do item da ap�lice inv�lido (n�o num�rico).")
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If Val(txtRenItem) = 0 Then
        Call gM46V999.gpGraLog(1, "FGE0070 - " & gM46V999.gPrefixo & _
                                  "N�mero do item da ap�lice inv�lido (zerado).")
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If Len(Trim(txtRenItem)) < 5 Then
        txtRenItem = Trim$(txtRenItem)
        If Len(txtRenItem) < 5 Then
            txtRenItem = String(5 - Len(txtRenItem), "0") & txtRenItem
        End If
    End If

    '4. Obt�m ap�lice/item.
    If gM46V999.gfAbrBasDados("P00APOL", gM46V999.gTipP00APOL, gM46V999.gEndP00APOL, lbdP00APOL, lMensagem) = False Then
        Call gM46V999.gpGraLog(1, lMensagem)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    'If gM46V999.gBaseLocal = True Then
    lIdeTabela = "TAB_APOL_LOC"
    'Else
    '    lIdeTabela = "GED.DBO.TAB_PED_LOC"
    'End If

    lSelect = "SELECT * FROM " & lIdeTabela & _
            " WHERE " & lIdeTabela & ".Num_Apol = " & txtRenApolice & _
            " AND " & lIdeTabela & ".Num_Item = " & txtRenItem

    If gM46V999.gfObtRegistro(lbdP00APOL, lSelect, rsAPOL, lMensagem) = False Then
        Call gM46V999.gpFecha2(lbdP00APOL)
        Call gM46V999.gpGraLog(1, lMensagem)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If rsAPOL.EOF = True Then
        'If gM46V999.gBaseLocal = True Then
        lIdeTabela = "TAB_APOL"
        'Else
        '    lIdeTabela = "GED.DBO.TAB_PED"
        'End If

        lSelect = "SELECT * from " & lIdeTabela & _
                " WHERE " & lIdeTabela & ".Num_Apol = " & txtRenApolice
        lSelect = lSelect & " and " & lIdeTabela & ".Num_Item = " & txtRenItem
        If gM46V999.gfObtRegistro(lbdP00APOL, lSelect, rsAPOL, lMensagem) = False Then
            Call gM46V999.gpFecha2(lbdP00APOL)
            Call gM46V999.gpGraLog(1, lMensagem)
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        If rsAPOL.EOF = True Then
            lMensagem = "FNL0008 - Ap�lice " & txtRenApolice & " n�o localizada."
        Else
            lMensagem = "FNL0009 - Item  " & txtRenItem & " n�o localizado na ap�lice " & _
                        txtRenApolice & "."
        End If

        'proposta eletr�nica
        If IsNumeric(rsAPOL("NUM_CGC_CPF")) Then
            mAntCNPJ = Format(Val(rsAPOL("NUM_CGC_CPF")), "00000000000000")
        Else
            mAntCNPJ = Right$(String(14, "0") & Trim(rsAPOL("NUM_CGC_CPF")), 14)
        End If
        mAntSeg = rsAPOL("NOM_SEGURADO")

        Call gM46V999.gpFecha2(lbdP00APOL)
        Call gM46V999.gpGraLog(1, lMensagem)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    gCarga = True
    frmT46V101A.txtRisTipLogradouro = gM46V999.gfTratarNulo(rsAPOL("Cod_Tip_Logr_Risco"), tpoString)
    frmT46V101A.txtRisNomLogradouro = gM46V999.gfTratarNulo(rsAPOL("Nom_Logr_Risco"), tpoString)
    frmT46V101A.txtRisComplemento = gM46V999.gfTratarNulo(rsAPOL("Dsc_Compl_Risco"), tpoString)
    frmT46V101A.txtRisBairro = gM46V999.gfTratarNulo(rsAPOL("Nom_Bairro_Risco"), tpoString)
    frmT46V101A.txtRisCidade = gM46V999.gfTratarNulo(rsAPOL("Nom_Cid_Risco"), tpoString)
    frmT46V101A.cboRisUF.ListIndex = -1
    If gM46V999.gfPreenchido(gM46V999.gfTratarNulo(rsAPOL("Sig_UF_Risco"), tpoString)) Then
        For lI = 0 To (frmT46V101A.cboRisUF.ListCount - 1)
            If gM46V999.gfTratarNulo(rsAPOL("Sig_UF_Risco"), tpoString) = frmT46V101A.cboRisUF.List(lI) Then
                frmT46V101A.cboRisUF.ListIndex = lI
                Exit For
            End If
        Next lI
    End If
    frmT46V101A.txtRisCEP = gM46V999.gfTratarNulo(rsAPOL("Num_CEP_Risco"), tpoNumerico)
    frmT46V101A.txtDescProgr = Format(gM46V999.gfTratarNulo(rsAPOL("Per_Desc_Progres"), TpoValor), "standard")
    ' frmT46V101A.txtISBas = Format(gM46V999.gfTratarNulo(rsAPOL("Per_IS_Vr"), TpoValor), "standard")
    frmT46V101A.txtCoefISVR = Format(gM46V999.gfTratarNulo(rsAPOL("Val_Coef_IS_Vr"), TpoValor), "standard")
    frmT46V101A.txtValoRisco = Format(gM46V999.gfTratarNulo(rsAPOL("Val_Vr"), TpoValor), "standard")
    'frmT46V101A.txtClassificacao.Tag = gM46V999.gfTratarNulo(rsAPOL("Cod_Clasf"), tpoNumerico) & gM46V999.gfTratarNulo(rsAPOL("Cod_Compl_Clasf"), tpoNumerico) & Format(gM46V999.gfTratarNulo(rsAPOL("Tab_Clasf"), tpoNumerico), "0000")
    'frmT46V101A.txtClassificacao = gM46V999.gfTratarNulo(rsAPOL("Cod_Clasf"), tpoNumerico) & "-" & gM46V999.gfTratarNulo(rsAPOL("Dsc_Clasf"), tpoString)
    frmT46V101A.txtClassificacao.Text = Format(gM46V999.gfTratarNulo(rsAPOL("Cod_Clasf"), tpoNumerico), "000 ") & "-" & gM46V999.gfTratarNulo(rsAPOL("Cod_Compl_Clasf"), tpoNumerico) & "-" & gM46V999.gfTratarNulo(rsAPOL("Dsc_Clasf"), tpoString)
    frmT46V101A.txtClassificacao.Tag = gM46V999.gfTratarNulo(rsAPOL("Cod_Clasf"), tpoNumerico) & String(4 - Len(gM46V999.gfTratarNulo(rsAPOL("Cod_Clasf"), tpoNumerico)), " ") & Format(gM46V999.gfTratarNulo(rsAPOL("Cod_Compl_Clasf"), tpoNumerico), "00") & Format(gM46V999.gfTratarNulo(rsAPOL("Tab_Clasf"), tpoNumerico), "0000")
    'Ficha 1000305
    If gM46V999.gCod_Ramo = e_CodRamo_Empresarial Then
        Select Case gM46V999.gfTratarNulo(rsAPOL("Cod_Constr"), tpoNumerico)
        Case e_CodContr_Superior
            frmT46V101A.cboConstrucao.ListIndex = 1
        Case e_CodContr_Inferior
            frmT46V101A.cboConstrucao.ListIndex = 2
        Case e_CodContr_Mista
            frmT46V101A.cboConstrucao.ListIndex = 3
        Case Else
            frmT46V101A.cboConstrucao.ListIndex = 1
        End Select
    Else
        If gM46V999.gfTratarNulo(rsAPOL("Cod_Constr"), tpoNumerico) = e_CodContr_Superior Then
            frmT46V101A.optContrucao(0).Value = True
        Else
            frmT46V101A.optContrucao(1).Value = True
        End If
    End If

    'Ficha 1000305
    '== Tipo Cobertura
    If gM46V999.gCod_Ramo = e_CodRamo_Empresarial Then
        Select Case gM46V999.gfTratarNulo(rsAPOL("Tip_Cobert"), tpoNumerico)
        Case e_TipCobert_PredioeConteudo
            frmT46V101A.cboTipCobert.ListIndex = 0
        Case e_TipCobert_Predio
            frmT46V101A.cboTipCobert.ListIndex = 1
        Case e_TipCobert_Conteudo
            frmT46V101A.cboTipCobert.ListIndex = 2
        Case Else
            frmT46V101A.cboTipCobert.ListIndex = 0
        End Select
    Else
        frmT46V101A.cboTipCobert.ListIndex = 0
    End If
    lBonus = gM46V999.gfTratarNulo(rsAPOL("Per_Bonus"), tpoNumerico)

    '    If optSinistroYas(0).Value = True Then 'Com sinistro
    '        lBonus = lBonus - 5
    '    Else
    '        lBonus = lBonus + 5
    '    End If
    If lBonus < 0 Then lBonus = 0

    'If pTipEmissao <> 300 Then
    '    If gM46V999.gCod_Ramo = e_CodRamo_Empresarial Then
    '        If gM46V999.gCodProduto >= 942 Then
    '            'Soma mais 5 da fidelidade.
    '           lBonus = lBonus + 5
    '       End If
    '       If lBonus > 30 Then lBonus = 30
    '    ElseIf gM46V999.gCod_Ramo = e_CodRamo_Residencial Then
    '        If gM46V999.gCodProduto >= 946 Then
    '            'Soma mais 5 da fidelidade.
    '            lBonus = lBonus + 5
    '            If lBonus > 30 Then lBonus = 30
    '        Else
    '            If lBonus > 25 Then lBonus = 25
    '       End If
    '    Else
    '        If lBonus > 25 Then lBonus = 25
    '    End If
    '
    '    'Ficha 901673 e 901772  'Ficha 1000099 - Planos 2131 a 2138 Corretor Schroeder
    '    If (gM46V111.gCod_Plano >= 2001 And gM46V111.gCod_Plano <= 2031) Or (gM46V111.gCod_Plano >= 2051 And gM46V111.gCod_Plano <= 2100) Or (gM46V111.gCod_Plano >= 2111 And gM46V111.gCod_Plano <= 2138) Then
    '        lBonus = 0
    '    End If
    'End If
    '
    'Ser� carregado no c�lculo
    frmT46V101A.txtDescBonus = lBonus

    frmT46V101A.txtQuaEmpregados = gM46V999.gfTratarNulo(rsAPOL("Qtd_Empreg"), tpoNumerico)

    'Seguran�a
    If gM46V999.gfTratarNulo(rsAPOL("Cod_Iden_Segur_01"), tpoString) = "S" Or _
       gM46V999.gfTratarNulo(rsAPOL("Cod_Iden_Segur_02"), tpoString) = "S" Or _
       gM46V999.gfTratarNulo(rsAPOL("Cod_Iden_Segur_03"), tpoString) = "S" Or _
       gM46V999.gfTratarNulo(rsAPOL("Cod_Iden_Segur_04"), tpoString) = "S" Or _
       gM46V999.gfTratarNulo(rsAPOL("Cod_Iden_Segur_05"), tpoString) = "S" Then
        frmT46V101A.optSeguranca(0).Value = True
        If gM46V999.gfTratarNulo(rsAPOL("Cod_Iden_Segur_01"), tpoString) = "S" Then
            frmT46V128A.chkSeguranca(0).Value = vbChecked
        Else
            frmT46V128A.chkSeguranca(0).Value = vbUnchecked
        End If
        If gM46V999.gfTratarNulo(rsAPOL("Cod_Iden_Segur_02"), tpoString) = "S" Then
            frmT46V128A.chkSeguranca(1).Value = vbChecked
        Else
            frmT46V128A.chkSeguranca(1).Value = vbUnchecked
        End If
        If gM46V999.gfTratarNulo(rsAPOL("Cod_Iden_Segur_03"), tpoString) = "S" Then
            frmT46V128A.chkSeguranca(2).Value = vbChecked
        Else
            frmT46V128A.chkSeguranca(2).Value = vbUnchecked
        End If
        If gM46V999.gfTratarNulo(rsAPOL("Cod_Iden_Segur_04"), tpoString) = "S" Then
            frmT46V128A.chkSeguranca(3).Value = vbChecked
        Else
            frmT46V128A.chkSeguranca(3).Value = vbUnchecked
        End If
        If gM46V999.gfTratarNulo(rsAPOL("Cod_Iden_Segur_05"), tpoString) = "S" Then
            frmT46V128A.chkSeguranca(4).Value = vbChecked
        Else
            frmT46V128A.chkSeguranca(4).Value = vbUnchecked
        End If
    Else
        frmT46V101A.optSeguranca(1).Value = True
        Unload frmT46V128A
    End If

    If gM46V999.gfTratarNulo(rsAPOL("Cod_Iden_Protec_01"), tpoString) = "S" Or _
       gM46V999.gfTratarNulo(rsAPOL("Cod_Iden_Protec_02"), tpoString) = "S" Or _
       gM46V999.gfTratarNulo(rsAPOL("Cod_Iden_Protec_03"), tpoString) = "S" Or _
       gM46V999.gfTratarNulo(rsAPOL("Cod_Iden_Protec_04"), tpoString) = "S" Then
        frmT46V101A.optProtecao(0).Value = True
        If gM46V999.gfTratarNulo(rsAPOL("Cod_Iden_Protec_01"), tpoString) = "S" Then
            frmT46V126A.chkProtecao(0).Value = vbChecked
        Else
            frmT46V126A.chkProtecao(0).Value = vbUnchecked
        End If
        If gM46V999.gfTratarNulo(rsAPOL("Cod_Iden_Protec_02"), tpoString) = "S" Then
            frmT46V126A.chkProtecao(1).Value = vbChecked
        Else
            frmT46V126A.chkProtecao(1).Value = vbUnchecked
        End If
        If gM46V999.gfTratarNulo(rsAPOL("Cod_Iden_Protec_03"), tpoString) = "S" Then
            frmT46V126A.chkProtecao(2).Value = vbChecked
        Else
            frmT46V126A.chkProtecao(2).Value = vbUnchecked
        End If
        If gM46V999.gfTratarNulo(rsAPOL("Cod_Iden_Protec_04"), tpoString) = "S" Then
            frmT46V126A.chkProtecao(3).Value = vbChecked
        Else
            frmT46V126A.chkProtecao(3).Value = vbUnchecked
        End If
    Else
        frmT46V101A.optProtecao(1).Value = True
        Unload frmT46V126A
    End If

    frmT46V101A.txtRisRegiao.Tag = Format(gM46V999.gfTratarNulo(rsAPOL("Cod_Regiao"), tpoNumerico), "00000") & Format(gM46V999.gfTratarNulo(rsAPOL("Cod_Produto"), tpoNumerico), "00000")
    gM46V111.gCod_Regiao = Val(Right(frmT46V101A.txtRisRegiao.Tag, 4))
    frmT46V101A.objstcPedLoc.Cod_Regiao = gM46V999.gfTratarNulo(rsAPOL("Cod_Regiao"), tpoNumerico)
    frmT46V101A.objstcPedLoc.Cod_Produto = gM46V999.gfTratarNulo(rsAPOL("Cod_Produto"), tpoNumerico)
    frmT46V101A.txtRisRegiao = frmT46V101A.objstcPedLoc.gDesc_Regiao

    'Clausulas
    'If gM46V999.gBaseLocal = True Then
    lIdeTabela = "TAB_APOL_CLAU"
    'Else
    '    lIdeTabela = "GED.DBO.TAB_PED_CLAU"
    'End If

    lSelect = " Select * from " & lIdeTabela & " Where Num_Apol = " & rsAPOL("Num_Apol")
    lSelect = lSelect & " and num_item = " & rsAPOL("Num_Item")

    If Not gM46V999.gfObtRegistro(lbdP00APOL, lSelect, rsAPOL, lMensagem) Then
        Exit Sub
    End If
    'Clausula do benefici�rio.
    If Not rsAPOL.EOF Then
        If rsAPOL("Cod_Clau") = 4 Or rsAPOL("Cod_Clau") = 1 Then
            frmT46V101A.optBeneficiario(0).Value = True
        Else
            frmT46V101A.optBeneficiario(1).Value = True
        End If
    End If

    'If gM46V999.gBaseLocal = True Then
    lIdeTabela = "TAB_APOL_COBERT"
    'Else
    '    lIdeTabela = "GED.DBO.TAB_PED_COBERT"
    'End If

    'Coberturas
    lSelect = " Select * from " & lIdeTabela & " Where Num_Apol = " & txtRenApolice
    lSelect = lSelect & " and num_item = " & txtRenItem

    If Not gM46V999.gfObtRegistro(lbdP00APOL, lSelect, rsAPOL, lMensagem) Then
        Exit Sub
    End If
    Set ColecaoCobert = New Collection
    Do While Not rsAPOL.EOF
        Set lobjstcCobert = New stcA46V704B
        Set lobjstcCobert.M46V999 = gM46V999

        With lobjstcCobert

            .Cod_Cobert = gM46V999.gfTratarNulo(rsAPOL("Cod_Cobert"), tpoNumerico)
            If .Cod_Cobert = 1 Then
                frmT46V101A.txtISBas = Format(gM46V999.gfTratarNulo(rsAPOL("Val_IS"), TpoValor), "standard")
            End If

            ' .NOSSO_NUMERO = gM46V999.gfTratarNulo(rsAPOL("NOsso_Numero"), tpoString)
            ' .Num_Ped = gM46V999.gfTratarNulo(rsAPOL("Num_Ped"), tpoNumerico)

            .Cod_Org = gM46V999.gfTratarNulo(rsAPOL("Cod_Org"), tpoNumerico)
            .Cod_Ramo = gM46V999.gfTratarNulo(rsAPOL("Cod_Ramo"), tpoNumerico)
            .Num_Item = gM46V999.gfTratarNulo(rsAPOL("Num_item"), tpoNumerico)
            .Tab_Cobert = gM46V999.gfTratarNulo(rsAPOL("Tab_Cobert"), tpoNumerico)
            .Val_IS = gM46V999.gfTratarNulo(rsAPOL("Val_IS"), tpoNumerico)
            '.Per_Taxa = gM46V999.gfTratarNulo(rsAPOL("Per_Taxa"), tpoNumerico)
            .Per_Desc_Coml = gM46V999.gfTratarNulo(rsAPOL("Per_Desc_Coml"), tpoNumerico)
            .Per_Desc_Exp = gM46V999.gfTratarNulo(rsAPOL("Per_Desc_Exp"), tpoNumerico)
            .VAL_PR = gM46V999.gfTratarNulo(rsAPOL("Val_PR"), tpoNumerico)
            .Val_Pr_Anu = gM46V999.gfTratarNulo(rsAPOL("Val_Pr_Anu"), tpoNumerico)
            .Val_Pr_Tarif = gM46V999.gfTratarNulo(rsAPOL("Val_Pr_Tarif"), tpoNumerico)
            .Val_Pr_Tarif_Anu = gM46V999.gfTratarNulo(rsAPOL("Val_Pr_Tarif_Anu"), tpoNumerico)
            .Val_Pr_Resg = gM46V999.gfTratarNulo(rsAPOL("Val_Pr_Resg"), tpoNumerico)
            .Val_Pr_Resg_Anu = gM46V999.gfTratarNulo(rsAPOL("Val_Pr_Resg_Anu"), tpoNumerico)
            .Per_Franq = gM46V999.gfTratarNulo(rsAPOL("Per_Franq"), tpoNumerico)
            .Val_Franq_Min = gM46V999.gfTratarNulo(rsAPOL("Val_Franq_Min"), tpoNumerico)
            .Val_Franq_Max = gM46V999.gfTratarNulo(rsAPOL("Val_Franq_Max"), tpoNumerico)
            '.Num_Meses_PI = gM46V999.gfTratarNulo(rsAPOL("Num_Meses_PI"), tpoNumerico)
            '.Val_Fator_Franq = gM46V999.gfTratarNulo(rsAPOL("Val_Fator_Franq"), TpoValor)
            '.Alterar_Cobert = gM46V999.gfTratarNulo(rsAPOL("Alterar_Cobert"), tpoString)
            '.Val_Risco_Cob = gM46V999.gfTratarNulo(rsAPOL("Val_Risco_Cob"), TpoValor)
        End With
        ColecaoCobert.Add lobjstcCobert, lobjstcCobert.Class_ID
        rsAPOL.MoveNext
    Loop
    'Cobertura
    Unload frmT46V117A
    If ColecaoCobert.Count > 0 Then
        'Call frmT46V117A.mpSelecionarCoberturas(ColecaoCobert)
        gM46V999.gChamaSelecionarCoberturas = True
    End If

    frmT46V112A.txtRenApolice.Text = txtRenApolice
    frmT46V112A.txtRenItem.Text = txtRenItem

    Set ColecaoClau = Nothing
    Set ColecaoCobert = Nothing
    gCarga = False
    Call gpHabSalvar
    gAltT46V112A = True

    '7. Reposiciona informa��es do item e retorna ao form de item.
    Me.Hide
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
    If flxConsulta.Rows = 1 Then
        MsgBox "N�o existem outros item para endosssar", vbExclamation
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Dim lobjstcPedLoc As stcA46V708B
    Dim lTotal As Double

    Screen.MousePointer = vbHourglass

    'Abro a base P0044000
    ' Set lobjstcPedido = New stcA46V702B
    ' Set lobjstcPedido.M46V999 = gM46V999

    'Preciso identificar os itens pertencentes a esta cota��o para n�o inclu�-las na pesquina
    mNumerosItens = ""
    'Monto uma string com os n�meros dos itens

    For Each lobjstcPedLoc In frmT46V101A.objstcPedido.ColecaoPedLoc
        If lobjstcPedLoc.Num_Apol_Anterior <> 0 Then
            mNumeroApolice = lobjstcPedLoc.Num_Apol_Anterior
        End If
        mNumerosItens = mNumerosItens & ", '" & Format(lobjstcPedLoc.Num_Item, "0000") & "'"
    Next

    '   mNumerosItens = Mid(mNumerosItens, 2)
    Call mpCarregarGrid(0)

    If cmdOk.Visible = True And cmdOk.Enabled = True Then
        cmdOk.SetFocus
    End If
    fraOpcao.Visible = False
    lblAviso.Visible = False


    Screen.MousePointer = vbDefault

End Sub


Private Sub mpCarregarGrid(pItem As Integer)

    Dim bdP00APOL As ADODB.Connection    'Base de dados: P00Apol.
    Dim rsP00APOL As clsYasRecordSet    'Registro: P00Apol.
    Dim lNomeTab As String
    Dim BDP00Tarifa As ADODB.Connection
    Dim rstTabClas As clsYasRecordSet
    Dim lSelect As String
    Dim lMensagem As String

    'Consulto todos os itens desta ap�lice que n�o est�o nesta cota��o

    If gM46V999.gPrograma = "D046O003" Then
        lNomeTab = "Syasportal.dbo.TAB_APOL"
    Else
        'If gM46V999.gBaseLocal = True Then
        lNomeTab = "TAB_APOL"
        'Else
        '    lNomeTab = "GED.DBO.TAB_PED"
        'End If
    End If
    If gM46V999.gfAbrBasDados("P00APOL", gM46V999.gTipP00APOL, gM46V999.gEndP00APOL, bdP00APOL, lMensagem) = False Then
        Exit Sub
    End If


    lSelect = "SELECT * FROM " & lNomeTab & "_LOC "
    lSelect = lSelect & " Where Num_Apol = " & mNumeroApolice
    lSelect = lSelect & " order by NUM_ENDO desc "
    If gM46V999.gfObtRegistro(bdP00APOL, lSelect, rsP00APOL, lMensagem) = False Then
        Call gM46V999.gpFecha2(bdP00APOL)
        Exit Sub
    End If

    If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, BDP00Tarifa, lMensagem) = False Then
        Call gM46V999.gpGraLog(1, lMensagem)
        Exit Sub
    End If



    Call mpFormatarGrid
    With flxConsulta
        While Not rsP00APOL.EOF
            If InStr(mNumerosItens, "'" & Format(rsP00APOL("Num_Item"), "0000") & "'") = 0 Then
                .Rows = .Rows + 1
                .Row = .Rows - 1

                lSelect = " Select distinct dsc_clasf from Tbm_Classificacao "
                lSelect = lSelect & " Where cod_Ramo  = " & rsP00APOL("Cod_Ramo")
                lSelect = lSelect & " and Cod_Tarif = " & rsP00APOL("Cod_Iden_Tarif")
                lSelect = lSelect & " and cod_Clasf = '" & rsP00APOL("Cod_Clasf") & "'"
                lSelect = lSelect & " and Cod_Compl_Clasf = '" & rsP00APOL("Cod_Compl_Clasf") & "'"
                lSelect = lSelect & " and  tab_clasf = " & rsP00APOL("tab_clasf")

                If gM46V999.gfObtRegistro(BDP00Tarifa, lSelect, rstTabClas, lMensagem) = False Then
                    Call gM46V999.gpGraLog(1, lMensagem)
                    Exit Sub
                End If

                .TextMatrix(.Rows - 1, 0) = rsP00APOL("Num_Item")   'N�mero do item.
                .TextMatrix(.Rows - 1, 1) = Trim(rstTabClas("dsc_clasf"))
                .TextMatrix(.Rows - 1, 2) = Trim(rsP00APOL("cod_tip_Logr_Risco")) & " " & Trim(rsP00APOL("Nom_Logr_Risco")) & " " & Trim(rsP00APOL("Dsc_compl_Risco"))
                .TextMatrix(.Rows - 1, 3) = Format(rsP00APOL("Num_Cep_Risco"), "00000-000")
            End If

            rsP00APOL.MoveNext

        Wend

        .ColSel = 0
        .Sort = 1
        .Col = 0    'N�mero do item.
        'd) Seleciona primeira linha.
        If .Rows > 1 Then
            .Row = 1
        End If
    End With

End Sub

Private Sub mpFormatarGrid()

    With flxConsulta
        .Rows = 1
        .Cols = 4
        .Row = 0
        .FontWidth = 3.35
        'a) Posiciona largura das colunas.
        .ColWidth(0) = 405
        .ColWidth(1) = 2505
        .ColWidth(2) = 2175
        .ColWidth(3) = 945
        'b) Posiciona altura da linha de cabe�alho.
        .RowHeight(0) = 220
        'c) Posiciona alinhamento das colunas.
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignCenterCenter
        'd) Posiciona t�tulos.
        .Col = 0
        .Text = "Item"
        .Col = 1
        .Text = "Classifica��o"
        .Col = 2
        .Text = "Logradouro"
        .Col = 3
        .Text = "CEP"
    End With

End Sub
