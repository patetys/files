VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmT46V212A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PERMISSÕES COTADOR WEB"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9765
   ControlBox      =   0   'False
   ForeColor       =   &H00808080&
   Icon            =   "T46V212A.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   9765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frameReanalise 
      Caption         =   "   Bloqueios e Limites   "
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      Begin VB.TextBox txtLimiteComissao 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   6000
         MaxLength       =   5
         TabIndex        =   3
         Text            =   "0"
         ToolTipText     =   "Percentual máximo entre 0 e 99%"
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   7920
         TabIndex        =   7
         Top             =   5400
         Width           =   1215
      End
      Begin VB.CommandButton cmdSalvar 
         Caption         =   "Salvar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5760
         TabIndex        =   5
         Top             =   5400
         Width           =   1815
      End
      Begin VB.CheckBox ChkBloqueiaWeb 
         Caption         =   "Bloquear Siscota web"
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2040
         TabIndex        =   2
         ToolTipText     =   "Bloqueia todos os campos"
         Top             =   360
         Width           =   1935
      End
      Begin VB.CheckBox chkComissao 
         Caption         =   "Comissão"
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   240
         TabIndex        =   1
         ToolTipText     =   "Obriga reanálise caso percentual exceda a comissão informada na cotação web"
         Top             =   360
         Width           =   1575
      End
      Begin MSComctlLib.ImageList imgLista 
         Left            =   8640
         Top             =   360
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "T46V212A.frx":000C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "T46V212A.frx":035E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "T46V212A.frx":06B0
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "T46V212A.frx":0A02
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "T46V212A.frx":0E1D
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flxCoberturaReanalise 
         Height          =   4215
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   7435
         _Version        =   393216
         RowHeightMin    =   100
         GridLinesFixed  =   1
      End
      Begin VB.Label lblLimiteComissaoPercentual 
         Caption         =   "%"
         Height          =   255
         Left            =   6720
         TabIndex        =   9
         Top             =   400
         Width           =   255
      End
      Begin VB.Label lblLimiteComissao 
         Caption         =   "Limite de Comissão"
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   4440
         TabIndex        =   8
         Top             =   405
         Width           =   1455
      End
      Begin VB.Label lblMsg 
         ForeColor       =   &H00000080&
         Height          =   495
         Left            =   600
         TabIndex        =   6
         Top             =   5280
         Width           =   4815
      End
   End
End
Attribute VB_Name = "frmT46V212A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ColecaoCobert As Collection
Dim flagPremioInformado As Boolean

Const strChecked = "þ"
Const strUnChecked = "q"

Public CalcularAoSalvar As Boolean

Private Sub ChkBloqueiaWeb_Click()
    If ChkBloqueiaWeb.value = 1 Then
        frmT46V101A.objstcPedido.objstcCtrlEmis.Ind_Bloq_Web = 1
    Else
        frmT46V101A.objstcPedido.objstcCtrlEmis.Ind_Bloq_Web = 0
    End If
End Sub

Private Sub chkComissao_Click()
    '' 20171213: Ajuste percentual maximo de comissao do corretor
    'If chkComissao.Value Then
    '    'If CInt(frmT46V101A.objstcPedido.objstcCtrlEmis.Per_Max_Comis_Corr) = 0 Then
    '    frmT46V101A.objstcPedido.objstcCtrlEmis.Per_Max_Comis_Corr = frmT46V101A.txtComissao.Text
    '    'End If
    'Else
    '    frmT46V101A.objstcPedido.objstcCtrlEmis.Per_Max_Comis_Corr = 0
    'End If
    
    frmT46V101A.objstcPedido.objstcCtrlEmis.RA_Comissao = chkComissao.Value
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    Dim lBanco As ADODB.Connection
    Dim lMensagem As String
    Dim lrs As clsYasRecordSet
    Dim lSelect As String
    Dim lRegAfetados As Long
    Dim lobjstc As stcA46V755B             'stcA46V755B - RamosDiversos..Tab_Ped_Cobert_Compl reduzida
    Dim lobjstcCob As stcA46V704B
    Dim i As Integer
    
    '' Atualiza Comissao
    lSelect = "update RamosDiversos..tab_ctrl_emis_compl set RA_Comissao = " & CStr(chkComissao.Value) & " where nosso_numero = '" & gM46V111.gComNosNumero & "' " & vbCrLf

    '' Atualiza RA_C4 (Coberturas que geram reanálise)
    For i = 1 To ColecaoCobert.Count
        Set lobjstc = ColecaoCobert(i)
        
        If IsNull(lobjstc.RA_C4) Then lobjstc.RA_C4 = 0
        If IsNull(lobjstc.VAL_IS_RA) Then lobjstc.VAL_IS_RA = 0
        If IsNull(lobjstc.TIP_ORIGEM_C4) Then lobjstc.TIP_ORIGEM_C4 = 0
        
        If frmT46V101A.objstcPedLoc.Num_Item = lobjstc.Num_Item Then
            For Each lobjstcCob In frmT46V101A.objstcPedLoc.ColecaoPedCobert
                If lobjstcCob.Cod_Cobert = lobjstc.Codigo Then
                    lobjstcCob.RA_C4 = lobjstc.RA_C4
                    lobjstcCob.VAL_IS_RA = lobjstc.VAL_IS_RA
                    lobjstcCob.TIP_ORIGEM_C4 = lobjstc.TIP_ORIGEM_C4
                End If
            Next lobjstcCob
        End If
        
        lSelect = lSelect & "update RamosDiversos..Tab_Ped_Cobert_Compl set "
        lSelect = lSelect & "TIP_ORIGEM_C4 = " & CStr(lobjstc.TIP_ORIGEM_C4) & ", "
        lSelect = lSelect & "RA_C4 = " & CStr(lobjstc.RA_C4) & ", "
        lSelect = lSelect & "VAL_IS_RA = " & gM46V999.gfFormatarCamposSQL(lobjstc.VAL_IS_RA)
        lSelect = lSelect & " where NOSSO_NUMERO = '" & gM46V111.gComNosNumero & "' "
        lSelect = lSelect & " and num_item=" & CStr(lobjstc.Num_Item)
        lSelect = lSelect & " and tab_cobert=" & CStr(lobjstc.Tab_Cobert)
        lSelect = lSelect & " and cod_cobert=" & CStr(lobjstc.Codigo) & vbCrLf
    Next i
    
    
    If gM46V999.gfAbrBasDados("RAMOSDIVERSOS", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lBanco, lMensagem) = False Then
        MsgBox lMensagem, vbCritical
        GoTo SAIDA
    End If
    
    If gM46V999.gfExeSQL(lBanco, lSelect, lMensagem, lRegAfetados) > 0 Then
        MsgBox lMensagem, vbCritical
        GoTo SAIDA
    End If
    
    lblMsg.Caption = "Executado com sucesso!"
    
    cmdCancelar.Caption = "Fechar"

SAIDA:
    GoSub FechaObj
    
    If CalcularAoSalvar Then
        Unload Me
        Call frmT46V101A.mnuEdiCalPremios_Click
    End If

    Exit Sub
FechaObj:
    Set lrs = Nothing
    Return
End Sub

Private Sub flxCoberturaReanalise_DblClick()
    'lblMsg.Caption = flxCoberturaReanalise.row & " | " & flxCoberturaReanalise.Col
End Sub

Private Sub flxCoberturaReanalise_SelChange()
    Call FlexGridBackColor(flxCoberturaReanalise.row, True)
End Sub

Private Sub flxCoberturaReanalise_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Integer
    Dim lobjstcLoc As stcA46V708B
    Dim lobjstcCob As stcA46V704B
    
    '' 2018-01-16: inibir se desabilitado
    If gM46V999.gDesabilitarForms Then
        Exit Sub
    End If
    
    i = flxCoberturaReanalise.row
    
    With flxCoberturaReanalise
        If flxCoberturaReanalise.MouseCol = 0 Then
            If flxCoberturaReanalise.MouseRow >= 1 Then
                If CBool(flxCoberturaReanalise.TextMatrix(i, 7)) = True Then
                    Exit Sub
                End If
    
                .Col = 0
                
                If .CellPicture = 0 Then
                    If CBool(.TextMatrix(i, 7)) = False Then
                        Set .CellPicture = imgLista.ListImages.Item(4).Picture 'picChecked.Picture
                    Else
                        Set .CellPicture = imgLista.ListImages.Item(5).Picture 'picChecked.Picture
                    End If
                    .CellPictureAlignment = flexAlignCenterCenter
                    ColecaoCobert(i).RA_C4 = 1
                Else
                    Set .CellPicture = Nothing
                    ColecaoCobert(i).RA_C4 = 0
                End If
            Else
                If .MouseRow = 0 And .MouseCol = 0 Then
                    Call TriggerCheckbox(.MouseRow, .MouseCol)
                End If
            End If
        End If
    End With
    
    'Call RefreshGridCoberturas
End Sub

Private Sub TriggerCheckbox(iRow As Integer, iCol As Integer)
    Dim Check As Boolean
    
    With flxCoberturaReanalise
        If .TextMatrix(iRow, iCol) = strUnChecked Then
            .TextMatrix(iRow, iCol) = strChecked
            Check = True
        Else
            .TextMatrix(iRow, iCol) = strUnChecked
            Check = False
        End If
    End With
    
    Call TriggerCobert(Check)
End Sub

Private Sub TriggerCobert(Check As Boolean)
    Dim lobjstc As stcA46V755B             'stcA46V755B - RamosDiversos..Tab_Ped_Cobert_Compl reduzida
    Dim lobjstcCob As stcA46V704B
    Dim i As Integer
    
    For i = 1 To ColecaoCobert.Count
        If Check Then
            ColecaoCobert(i).RA_C4 = 1
        Else
            ColecaoCobert(i).RA_C4 = 0
        End If
        
        ColecaoCobert(i).TIP_ORIGEM_C4 = 1
                
        If IsNull(ColecaoCobert(i).RA_C4) Then ColecaoCobert(i).RA_C4 = 0
        If IsNull(ColecaoCobert(i).VAL_IS_RA) Then ColecaoCobert(i).VAL_IS_RA = 0
        If IsNull(ColecaoCobert(i).TIP_ORIGEM_C4) Then ColecaoCobert(i).TIP_ORIGEM_C4 = 0
        
        If frmT46V101A.objstcPedLoc.Num_Item = ColecaoCobert(i).Num_Item Then
            For Each lobjstcCob In frmT46V101A.objstcPedLoc.ColecaoPedCobert
                If lobjstcCob.Cod_Cobert = ColecaoCobert(i).Codigo Then
                    'lobjstcCob.RA_C4 = ColecaoCobert(i).RA_C4
                    'lobjstcCob.TIP_ORIGEM_C4 = ColecaoCobert(i).TIP_ORIGEM_C4
                    ColecaoCobert(i).VAL_IS_RA = lobjstcCob.VAL_IS_RA
                End If
            Next lobjstcCob
        End If
    Next i
    
    '' Refresh
    Call RefreshGridCoberturas
End Sub

Private Sub FlexGridBackColor(row As Integer, STATUS As Boolean)
    Dim r, c As Integer
        
    For r = 1 To flxCoberturaReanalise.Rows - 1
        For c = 1 To flxCoberturaReanalise.Cols - 1
            With flxCoberturaReanalise
                .row = r
                .Col = c
                If (STATUS And row = r) Then
                    If CBool(flxCoberturaReanalise.TextMatrix(r, 7)) = False Then
                        .CellBackColor = &HFFFF&        '' Selecionado
                    Else
                        .CellBackColor = &HC0C0C0       '' Selecionado
                    End If
                Else
                    .CellBackColor = &H8000000F     '' Limpar
                End If
            End With
        Next c
    Next r
End Sub

Private Sub Form_Activate()
    If CalcularAoSalvar Then
        cmdSalvar.Caption = "Salvar e Calcular"
    End If
    
    If gM46V999.gDesabilitarForms Then
        cmdSalvar.Enabled = False
        chkComissao.Enabled = False
        ChkBloqueiaWeb.Enabled = False
        txtLimiteComissao.Enabled = False
        flxCoberturaReanalise.Enabled = False
    End If
    
    Call mpPrepararGrid
End Sub

Private Sub Form_Load()
    If frmT46V101A.cboPlano.ListIndex = -1 Then
        GoTo Fim
    End If
    
    '' Check na carga
    If frmT46V101A.objstcPedido.objstcCtrlEmis.Ind_Bloq_Web = 1 Then
        ChkBloqueiaWeb.Value = 1
    End If

    '' Limite Comissao
    txtLimiteComissao.Text = CStr(frmT46V101A.objstcPedido.objstcCtrlEmis.Per_Max_Comis_Corr)
    
    '' Prepara grid
    Call mpPrepararGrid
    
    Call gM46V999.gpCenForm(Me)

Fim:
    ' Fim
End Sub

Private Sub mpPrepararGrid()
    With flxCoberturaReanalise
        .Rows = 1
        .Cols = 8
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Item"
        .TextMatrix(0, 2) = "Cód"
        If gM46V999.gCod_Ramo = e_CodRamo_Roubos Then
            .TextMatrix(0, 3) = "Bem coberto"
        Else
            .TextMatrix(0, 3) = "Cobertura"
        End If
        .TextMatrix(0, 4) = "Tab_Cobert"
        .TextMatrix(0, 5) = "VAL_IS"
        .TextMatrix(0, 6) = "VAL_PR"

        .row = 0
        .ColWidth(0) = 400
        .ColWidth(1) = 500
        .ColWidth(2) = 500
        .ColWidth(3) = 6000
        .ColWidth(4) = 0
        .ColWidth(5) = 1200
        .ColWidth(6) = 0
        .ColWidth(7) = 0
        
        .row = 0
        .Col = 0
        .CellFontName = "Wingdings"
        .CellFontSize = 14
        .CellAlignment = flexAlignCenterCenter
        .Text = strUnChecked
    End With

    '' Carrega colecao
    Call mpCarregarGridCoberturas
    
    'Exibe
    Call RefreshGridCoberturas
End Sub

Public Sub mpCarregarGridCoberturas()
    If frmT46V101A.cboPlano.ListIndex <> -1 Then
        'Carrega Coleção de Coberturas
        Call mpCarregarColecaoCobert
    End If
End Sub

Private Sub mpCarregarColecaoCobert()
    Dim lobjstc As stcA46V755B                     'Objeto da classe de estrutura stcA46V753A // stcA46V101A.
    Dim lbdTAB_Cobert As ADODB.Connection                     'Arquivo: P00Tarifa.
    Dim lMensagem As String                     'Mensagem.
    Dim lrsTab_Cobert As clsYasRecordSet                     'Registro : TAB_COBERT.
    Dim lPredioConteudoVerbasSeparadas As Boolean     'Prédio e conteúdo com verbas separadas.
    Dim lSelect As String
    Dim lCoberturaRestrita As Integer

    Set ColecaoCobert = New Collection

    If (frmT46V101A.cboTipCobert.ListIndex = e_TipCobert_PredioeConteudo) And (frmT46V101A.optVerbasSeparadas(0).value = True) Then
        lPredioConteudoVerbasSeparadas = True
    Else
        lPredioConteudoVerbasSeparadas = False
    End If

    'Nova tarifa 1092 / 1093
    If val(frmT46V102A.txtCorretor(1).Tag) = 908624 And frmT46V101A.objstcPedido.Cod_Produto >= 1092 Then
        lCoberturaRestrita = True
    Else
        lCoberturaRestrita = False
    End If

    If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTAB_Cobert, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If


    lSelect = "SELECT TOP 5000 TPC.NUM_ITEM, TC.COD_COBERT, TC.DSC_COBERT, TC.TAB_COBERT, TPC.VAL_IS, TPC.VAL_PR, TPCC.PremioInformado, TPCC.RA_C4, TPCC.VAL_IS_RA, TPCC.TIP_ORIGEM_C4 " & vbCrLf
    lSelect = lSelect & "FROM RamosDiversos..TAB_PED_COBERT TPC (NOLOCK) " & vbCrLf
    lSelect = lSelect & "INNER JOIN RamosDiversos..Tab_Ped_Cobert_Compl TPCC (NOLOCK) on TPCC.NOSSO_NUMERO = TPC.NOSSO_NUMERO and TPCC.NUM_ITEM = TPC.NUM_ITEM and TPCC.COD_COBERT = TPC.COD_COBERT " & vbCrLf
    lSelect = lSelect & "INNER JOIN RamosDiversos..TAB_COBERT TC (NOLOCK) ON TPC.COD_RAMO=TC.COD_RAMO AND TPC.TAB_COBERT = TC.TAB_COBERT AND TPC.COD_COBERT = TC.COD_COBERT " & vbCrLf
    lSelect = lSelect & "WHERE TPC.NOSSO_NUMERO = '" & gM46V111.gComNosNumero & "'" & vbCrLf
    lSelect = lSelect & "AND " & gM46V999.gCodProduto & " BETWEEN COD_TARIF_INI AND COD_TARIF_TER" & vbCrLf
    lSelect = lSelect & "AND TC.COD_COBERT <> 1 " & vbCrLf

    If lPredioConteudoVerbasSeparadas = False Then
        lSelect = lSelect & " and TC.COD_GRUPO_COBERT in (1, 2, 3, 4)" & vbCrLf
    Else
        lSelect = lSelect & " and TC.COD_GRUPO_COBERT in (1, 2, 3, 4, 5)" & vbCrLf
    End If
    
    'Nova tarifa 1092 / 1093
    If lCoberturaRestrita Then
        lSelect = lSelect & " AND (cod_restrito IN(1,0,2)  or cod_restrito is null)" & vbCrLf
    Else
        'Residencial Supremo
        If frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex) = 30 Then
            lSelect = lSelect & " AND (cod_restrito IN(0,2)  or cod_restrito is null)" & vbCrLf
        Else
            lSelect = lSelect & " AND TC.COD_COBERT <> 1215" & vbCrLf
        End If
    End If

    If gM46V999.gCodProduto >= 1122 Then '11.22
        gM46V999.gNum_Apolice = frmT46V101A.objstcPedido.Num_Apol_An
        If Not (frmT46V101A.objstcPedLoc.Cod_DizerItem >= e_CodDizerInclusao And gM46V999.gTarifa_Apolice > 0 And gM46V999.gTarifa_Apolice < 1122) Then
            lSelect = lSelect & " AND NOT (TC.COD_RAMO = 114 AND TC.COD_COBERT = 9636 AND TC.TAB_COBERT = 11)" & vbCrLf
        End If
    End If

    lSelect = lSelect & "ORDER BY TPC.NUM_ITEM, NUM_ORDEM_COBERT, TC.Cod_Cobert"


    If gM46V999.gfObtRegistro(lbdTAB_Cobert, lSelect, lrsTab_Cobert, lMensagem) = False Then
        Call gM46V999.gpFecha2(lbdTAB_Cobert)
        Call gM46V999.gpGraLog(2, lMensagem)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Set ColecaoCobert = New Collection

    Do While Not lrsTab_Cobert.EOF
        Set lobjstc = New stcA46V755B
        
        lobjstc.Num_Item = Trim(lrsTab_Cobert("NUM_ITEM"))
        lobjstc.Codigo = lrsTab_Cobert("COD_COBERT")
        lobjstc.Descricao = Trim(lrsTab_Cobert("DSC_COBERT"))
        lobjstc.Tab_Cobert = lrsTab_Cobert("TAB_COBERT")
        lobjstc.Val_IS = lrsTab_Cobert("VAL_IS")
        lobjstc.Val_Pr = lrsTab_Cobert("VAL_PR")
        
        If lrsTab_Cobert("RA_C4") = 1 Then
            lobjstc.RA_C4 = 1
        Else
            lobjstc.RA_C4 = 0
        End If
        
        If IsNull(lrsTab_Cobert("VAL_IS")) Then
            lobjstc.VAL_IS_RA = 0
        Else
            lobjstc.VAL_IS_RA = lrsTab_Cobert("VAL_IS")
        End If
        
        If IsNull(lrsTab_Cobert("TIP_ORIGEM_C4")) Then
            lobjstc.TIP_ORIGEM_C4 = 0
        Else
            lobjstc.TIP_ORIGEM_C4 = lrsTab_Cobert("TIP_ORIGEM_C4")
        End If
        
        If IsNull(lrsTab_Cobert("PremioInformado")) Then
            lobjstc.PremioInformado = False
        ElseIf lrsTab_Cobert("PremioInformado") = 0 Then
            lobjstc.PremioInformado = False
        Else
            lobjstc.PremioInformado = True
            lobjstc.RA_C4 = 1
            ''lobjstc.TIP_ORIGEM_C4 = 1
        End If
        
        ColecaoCobert.Add lobjstc
        lrsTab_Cobert.MoveNext
    Loop
End Sub


Private Sub RefreshGridCoberturas()
    Dim lLinha As Long
    Dim lobjstc As stcA46V755B             'Objeto da consulta parcial de coberturas
    Dim flagPremioInfo As Boolean

    lLinha = 1
    flagPremioInformado = False
    
    For Each lobjstc In ColecaoCobert
        flxCoberturaReanalise.Rows = lLinha + 1

        If lobjstc.PremioInformado Then
            flagPremioInformado = True
            flagPremioInfo = True
        Else
            flagPremioInfo = False
        End If

        With flxCoberturaReanalise
            .row = lLinha
            If lobjstc.RA_C4 > 0 Then
                .Col = 0
                If flagPremioInfo Then
                    Set .CellPicture = imgLista.ListImages.Item(5).Picture
                Else
                    Set .CellPicture = imgLista.ListImages.Item(4).Picture
                End If
                .CellPictureAlignment = flexAlignCenterCenter
            Else
                .Col = 0
                Set .CellPicture = Nothing
            End If
            
            .TextMatrix(lLinha, 1) = lobjstc.Num_Item
            .TextMatrix(lLinha, 2) = lobjstc.Codigo
            .TextMatrix(lLinha, 3) = lobjstc.Descricao
            .TextMatrix(lLinha, 4) = lobjstc.Tab_Cobert
            .TextMatrix(lLinha, 5) = FormatNumber(lobjstc.Val_IS, 2, True, True, True)
            .TextMatrix(lLinha, 6) = FormatNumber(lobjstc.Val_Pr, 2, True, True, True)
            .TextMatrix(lLinha, 7) = lobjstc.PremioInformado
        End With
        
        lLinha = lLinha + 1
    Next lobjstc
    
    '' 2017-12-13
    If frmT46V101A.objstcPedido.objstcCtrlEmis.RA_Comissao > 0 Then
        chkComissao.Value = 1
    Else
        chkComissao.Value = 0
    End If
    
    '' 2017-12-21
    If flagPremioInformado Then
        chkComissao.Value = 1
        chkComissao.Enabled = False
    Else
        chkComissao.Enabled = True
    End If

End Sub

Private Sub txtLimiteComissao_Change()
    Dim Per_Max_Comis_Corr As String
    Dim lPer_Max_Comis_Corr As Double

    ' Verificar se deseja ter um limite de comissão
    Per_Max_Comis_Corr = txtLimiteComissao.Text
    
    If Per_Max_Comis_Corr = "" Then
        MsgBox "Limite de comissão removido.", vbInformation, "Limite não informado"
        lPer_Max_Comis_Corr = 0
        txtLimiteComissao.Text = "0"
    ElseIf IsNumeric(Per_Max_Comis_Corr) Then
        lPer_Max_Comis_Corr = CDbl(Per_Max_Comis_Corr)
        
        If lPer_Max_Comis_Corr < 0 Or lPer_Max_Comis_Corr > 99 Then
            MsgBox "Valor inválido. Valores permitidos são entre 0 e 99", vbExclamation, "Valor inválido"
            lPer_Max_Comis_Corr = 0
            txtLimiteComissao.Text = "0"
        End If
    Else
        txtLimiteComissao.Text = "0"
    End If
    
    frmT46V101A.objstcPedido.objstcCtrlEmis.Per_Max_Comis_Corr = lPer_Max_Comis_Corr
End Sub

Private Sub txtLimiteComissao_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode < 48 Or KeyCode > 57 Then
        DoEvents
    End If
End Sub
