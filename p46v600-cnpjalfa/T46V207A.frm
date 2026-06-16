VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmT46V207A 
   Caption         =   "Solicita��o de inspe��o"
   ClientHeight    =   5220
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13440
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5220
   ScaleWidth      =   13440
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList imgLista 
      Left            =   12480
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V207A.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V207A.frx":0352
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V207A.frx":06A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V207A.frx":09F6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13440
      _ExtentX        =   23707
      _ExtentY        =   1164
      ButtonWidth     =   1535
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgLista"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Consultar"
            ImageIndex      =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Editar"
            Key             =   "EditarCobert"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair"
            Key             =   "Sair"
            ImageIndex      =   1
         EndProperty
      EndProperty
      Begin VB.Frame fraCondicaoInspGeral 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Condi��o de Inspe��o para todos os itens:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   585
         Left            =   8950
         TabIndex        =   4
         Top             =   0
         Width           =   4395
         Begin Threed.SSOption optCondicaoInspecao 
            Height          =   315
            Index           =   0
            Left            =   165
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   195
            Width           =   1545
            _Version        =   65536
            _ExtentX        =   2725
            _ExtentY        =   556
            _StockProps     =   78
            Caption         =   "Conforme Normas"
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   1
         End
         Begin Threed.SSOption optCondicaoInspecao 
            Height          =   315
            Index           =   1
            Left            =   1890
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   195
            Width           =   1185
            _Version        =   65536
            _ExtentX        =   2090
            _ExtentY        =   556
            _StockProps     =   78
            Caption         =   "Dispensada"
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   1
         End
         Begin Threed.SSOption optCondicaoInspecao 
            Height          =   315
            Index           =   2
            Left            =   3210
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   195
            Width           =   1125
            _Version        =   65536
            _ExtentX        =   1984
            _ExtentY        =   556
            _StockProps     =   78
            Caption         =   "Obrigat�ria"
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   1
         End
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   345
      Left            =   30
      TabIndex        =   1
      Top             =   6300
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   2
      Top             =   4785
      Width           =   13440
      _ExtentX        =   23707
      _ExtentY        =   767
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid flxSolicInspecao 
      Height          =   4005
      Left            =   60
      TabIndex        =   3
      Top             =   750
      Width           =   13320
      _ExtentX        =   23495
      _ExtentY        =   7064
      _Version        =   393216
      Rows            =   5
      Cols            =   4
      FixedCols       =   0
      BackColor       =   -2147483637
      ForeColor       =   8388608
      BackColorFixed  =   -2147483636
      ForeColorFixed  =   16777215
      ForeColorSel    =   12648447
      BackColorBkg    =   -2147483633
      GridColorFixed  =   4210752
      AllowBigSelection=   0   'False
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
   End
End
Attribute VB_Name = "frmT46V207A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private mobjstcInspecao As stcInspecao
Private mColecaoInspecao As Collection
Dim mobjstcPed As stcA46V702B
Dim mobjstcPedLoc As stcA46V708B
Const strChecked = "�"
Const strUnChecked = "q"

Private Sub ConfigGrid()
    With flxSolicInspecao
        .Cols = 11
        .Rows = 1
                       
        .Col = 0
        .TextMatrix(0, 0) = "Item"
        .ColWidth(0) = 400
        .TextMatrix(0, 1) = "Conforme Normas"
        .ColWidth(1) = 1400
        .TextMatrix(0, 2) = "Dispensada"
        .ColWidth(2) = 1000
        .TextMatrix(0, 3) = "Obrigat�ria"
        .ColWidth(3) = 1000
        .TextMatrix(0, 4) = "Programada"
        .ColWidth(4) = 1000
        .TextMatrix(0, 5) = "Inspe��o"
        .ColWidth(5) = 800
        .TextMatrix(0, 6) = "Local de Risco"
        .ColWidth(6) = 2500
        .TextMatrix(0, 7) = "Complemento"
        .ColWidth(7) = 1200
        .TextMatrix(0, 8) = "Valor de Risco"
        .ColWidth(8) = 1500
        .TextMatrix(0, 9) = "J� Solicitado"
        .ColWidth(9) = 1000
        .TextMatrix(0, 10) = "Insp obrigat�ria"
        .ColWidth(10) = 1200
        .Col = 0
    End With
End Sub

Public Sub Carregar_Form(ByVal pobjstcPed As stcA46V702B)
    Call ConfigGrid
    Set mobjstcPed = pobjstcPed

    With flxSolicInspecao
        For Each mobjstcPedLoc In mobjstcPed.ColecaoPedLoc
            .Rows = .Rows + 1
            Call Carregar_Inspecao_Linha(mobjstcPedLoc.Num_Item, mobjstcPedLoc.Ind_StatusInspecao, _
            mobjstcPedLoc.Num_Insp, mobjstcPedLoc.Nom_Logr_Risco, mobjstcPedLoc.objstcPedInsp.Num_Logr_Risco, _
            mobjstcPedLoc.Dsc_Compl_Risco, mobjstcPedLoc.Val_Vr, mobjstcPedLoc.NOSSO_NUMERO, mobjstcPedLoc.Insp_Obr)
        Next
    End With
End Sub

Public Sub Carregar_Inspecao_Linha(ByVal pNum_Item As Long, ByVal pInd_StatusInspecao As Integer, _
                                   ByVal pNum_Insp As Long, ByVal pNom_Logr_Risco As String, ByVal pNum_Logr_Risco As String, _
                                   ByVal pDsc_Compl_Risco As String, ByVal pVal_Vr As Double, ByVal pNosso_Numero As String, ByVal pInsp_Obr As Boolean)
    
    Dim lSql As String
    Dim lrst As clsYasRecordSet
    Dim lBD As ADODB.Connection
    Dim lMensagem As String
    Dim lI As Integer
    Dim i As Integer
    
    On Error GoTo Erro
    
    With flxSolicInspecao
        .Row = GetRow_ByNumItem(pNum_Item)
        .Col = 0
        .TextMatrix(.Row, 0) = pNum_Item
        .CellAlignment = flexAlignCenterCenter
            
        For i = 1 To 4
            .Col = i
            .CellFontName = "Wingdings"
            .CellFontSize = 14
            .CellAlignment = flexAlignCenterCenter
            .Text = strUnChecked
        Next i
        
        If gM46V111.gCod_plano > 100 Then
            .Col = 1
        Else
            Select Case pInd_StatusInspecao
                Case e_StatusInspecao_ConfNormas
                    .TextMatrix(.Row, 1) = strChecked
                Case e_StatusInspecao_Dispensada
                    .TextMatrix(.Row, 2) = strChecked
                Case e_StatusInspecao_obrigatoria
                    .TextMatrix(.Row, 3) = strChecked
                Case e_StatusInspecao_programada
                    .TextMatrix(.Row, 4) = strChecked
            End Select
        End If
        .TextMatrix(.Row, 5) = pNum_Insp
        .TextMatrix(.Row, 6) = pNom_Logr_Risco & "," & pNum_Logr_Risco
        .TextMatrix(.Row, 7) = pDsc_Compl_Risco
        .TextMatrix(.Row, 8) = "R$ " & Format(pVal_Vr, "standard")
        .TextMatrix(.Row, 9) = "N�o"
        
        If pInsp_Obr = True Then
            .TextMatrix(.Row, 10) = "Sim"
        Else
            .TextMatrix(.Row, 10) = "N�o"
        End If
    End With
    
    Exit Sub
Erro:
    Call gM46V999.gpGraLog(0, Me.Name + "_Carregar_Inspecao_Linha ERRO - " & Err & " " & Error)
End Sub

Private Function GetRow_ByNumItem(ByVal pNum_Item As Long) As Integer
    Dim lrow As Integer
    
    On Error GoTo Erro
    
    lrow = 1
    Do While lrow < flxSolicInspecao.Rows
        If flxSolicInspecao.TextMatrix(lrow, 0) = "" Then
            GetRow_ByNumItem = lrow
            Exit Do
        End If
        If flxSolicInspecao.TextMatrix(lrow, 0) = pNum_Item Then
            GetRow_ByNumItem = lrow
            Exit Do
        End If
        lrow = lrow + 1
    Loop
    
    Exit Function
Erro:
    GetRow_ByNumItem = 0
    Call gM46V999.gpGraLog(0, Me.Name + "_GetRow_ByNumItem - " & Err & " " & Error)
End Function

Private Function mfBuscaInspecao() As Boolean
    Dim lbdTab_Insp As ADODB.Connection      'Arquivo: P00Multi.
    Dim lMensagem As String                  'Mensagem.
    Dim lrsTab_Insp As clsYasRecordSet       'Registro : Tbm_Coef_PI.
    Dim lSelect As String
    Dim lDataParam As String
    Dim lIndice As String
    On Error GoTo Erro
    mfBuscaInspecao = False
    
    If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTab_Insp, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Function
    End If
    
    lIndice = 1
    lDataParam = gM46V999.gfForData(1, DateAdd("YYYY", -5, Date))
    Set mColecaoInspecao = New Collection
    For Each mobjstcPedLoc In mobjstcPed.ColecaoPedLoc
        lSelect = " Select top 20 num_insp, num_cep, tip_logr, nom_logr, num_logr, dsc_compl, dsc_cid, dsc_bairro, uf, cod_status, dat_derie, cod_user_cad, cod_clasf, cod_compl_clasf from inspecao..tab_insp (nolock) "
        lSelect = lSelect & " Where dat_cad >= " & lDataParam
        lSelect = lSelect & " and num_cep = " & mobjstcPedLoc.Num_Cep_Risco
        lSelect = lSelect & " and num_cnpj_cpf = '" & Trim(mobjstcPed.Num_Cgc_Cpf) & "'"
        lSelect = lSelect & " and cod_ramo = " & mobjstcPed.Cod_Ramo
        lSelect = lSelect & " order by num_insp desc "
    
        If gM46V999.gfObtRegistro(lbdTab_Insp, lSelect, lrsTab_Insp, lMensagem) = False Then
            Call gM46V999.gpFecha2(lbdTab_Insp)
            Call gM46V999.gpGraLog(2, lMensagem)
            Screen.MousePointer = vbDefault
            Exit Function
        End If

        Do While Not lrsTab_Insp.EOF
            Set mobjstcInspecao = New stcInspecao
            With mobjstcInspecao
                .Num_Ind = lIndice
                .Num_Item = mobjstcPedLoc.Num_Item
                .Num_Insp = lrsTab_Insp!Num_Insp
                .Num_CEP = lrsTab_Insp!Num_CEP
                .tip_logr = lrsTab_Insp!tip_logr
                .Nom_Logr = lrsTab_Insp!Nom_Logr
                .Num_Logr = lrsTab_Insp!Num_Logr
                .dsc_compl = lrsTab_Insp!dsc_compl
                .dsc_cid = lrsTab_Insp!dsc_cid
                .dsc_bairro = lrsTab_Insp!dsc_bairro
                .uf = lrsTab_Insp!uf
                .cod_status = lrsTab_Insp!cod_status
                .dat_derie = lrsTab_Insp!dat_derie
                .cod_user_cad = lrsTab_Insp!cod_user_cad
                .Cod_Clasf = lrsTab_Insp!Cod_Clasf
                .Cod_Compl_Clasf = lrsTab_Insp!Cod_Compl_Clasf
                 mColecaoInspecao.Add mobjstcInspecao
            End With
            lrsTab_Insp.MoveNext
        Loop
        lIndice = lIndice + 1
    Next
    If mobjstcInspecao Is Nothing Then
        MsgBox "N�o foram localizadas Inspe��es anteriores"
    Else
        MsgBox "Consulta realizada com sucesso!"
    End If
    mfBuscaInspecao = True
    Call gM46V999.gpFecha2(lbdTab_Insp)
    Exit Function
Erro:
    MsgBox Err.Number & " - " & Err.Description
End Function

Private Sub Salvar_Item()
    Dim PedLoc As stcA46V708B
    Dim lobjPedLoc As clsA46V708A
    Dim lMsgErro As String
    
    On Error GoTo Erro
    Screen.MousePointer = vbHourglass
    LockWindowUpdate flxSolicInspecao.hWnd
    
    Set lobjPedLoc = New clsA46V708A 'stcA46V708A
    Set lobjPedLoc.M46V999 = gM46V999
    
    For Each PedLoc In mobjstcPed.ColecaoPedLoc
        If val(frmT46V101A.txtNumItem.Text) = PedLoc.Num_Item Then
            Select Case PedLoc.Ind_StatusInspecao
                Case e_StatusInspecao_ConfNormas
                    frmT46V101A.optCondicaoInspecao(e_StatusInspecao_ConfNormas).Value = True
                Case e_StatusInspecao_Dispensada
                    frmT46V101A.optCondicaoInspecao(e_StatusInspecao_Dispensada).Value = True
                Case e_StatusInspecao_obrigatoria
                    frmT46V101A.optCondicaoInspecao(e_StatusInspecao_obrigatoria).Value = True
                Case e_StatusInspecao_programada
                    frmT46V101A.optCondicaoInspecao(e_StatusInspecao_programada).Value = True
            End Select
            frmT46V101A.txtNumInsp.Text = PedLoc.Num_Insp
        Else
'                .objstcPedObs.Observacao = Replace(frmT46V101A.Nova_Observacao(.Ind_StatusInspecao, .num_insp, lobjPedLoc), vbCrLf, "&%")
        End If
        If Not lobjPedLoc.mfAlterar(PedLoc, "", "", False) Then
            lMsgErro = lobjPedLoc.UltimoErro
            GoTo Erro
        End If
    Next PedLoc
    
    GoTo Fim:
Erro:
    If Err.Number <> 0 Then
        lMsgErro = lMsgErro & " " & Err.Number & " - " & Err.Description
    End If
    MsgBox lMsgErro
Fim:
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
    Dim PlanosRegulares As Boolean
    PlanosRegulares = Not mobjstcPed.getPlano.EhPlaninho
    If mobjstcPed.Cod_Estip = 10 Or mobjstcPed.Cod_Estip = 11 Or mobjstcPed.Cod_Estip = 20 Or mobjstcPed.Cod_Estip = 30 Then
        PlanosRegulares = True
    End If

    flxSolicInspecao.Enabled = Not (gM46V999.gDesabFormsCotacao Or PlanosRegulares = False) And gM46V999.gTipoAcesso <> DIGITACAO
    Toolbar.Enabled = Not (gM46V999.gDesabFormsCotacao Or PlanosRegulares = False) And gM46V999.gTipoAcesso <> DIGITACAO
    flxSolicInspecao.Enabled = flxSolicInspecao.Enabled And Not (mobjstcPed.Dat_Pri_Calculo <> 0 And gM46V999.gTipoAcesso = CALCULO And gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda And (UCase(gM46V999.gSerDIVERSOS) Like "*MIX*" = False) And Not UCase(gM46V999.gNomUsuMainFrame) Like "DSIS*")
End Sub

Private Sub Form_Load()
    Call gM46V999.gpCenForm(Me)
End Sub

Private Sub flxSolicInspecao_DblClick()
    If gM46V111.gCod_plano < 100 Then
        Call frmT46V208A.CarregarForm(mobjstcPed.ColecaoPedLoc, mColecaoInspecao, flxSolicInspecao.Row)
        Call Carregar_Form(mobjstcPed)
    End If
End Sub

Private Sub optCondicaoInspecao_Click(index As Integer, Value As Integer)
Dim PedLoc As stcA46V708B

    If MsgBox("Deseja selecionar esta op��o para todos os itens?", vbQuestion + vbYesNo) = vbYes Then
        Screen.MousePointer = vbHourglass
        For Each PedLoc In mobjstcPed.ColecaoPedLoc
            PedLoc.Ind_StatusInspecao = index
            PedLoc.Num_Insp = 0
            
            If index = e_StatusInspecao_obrigatoria Then
                PedLoc.Insp_Obr = True
            End If
            If index = e_StatusInspecao_Dispensada Then
                PedLoc.Insp_Obr = False
            End If
        Next PedLoc
        
        Call Carregar_Form(mobjstcPed)
'        Call ConfigGrid
'        With flxSolicInspecao
'            For Each PedLoc In mobjstcPed.ColecaoPedLoc
'                .Rows = .Rows + 1
'                Call Carregar_Inspecao_Linha(PedLoc.Num_Item, PedLoc.Ind_StatusInspecao, _
'                PedLoc.num_insp, PedLoc.Nom_Logr_Risco, PedLoc.objstcPedInsp.Num_Logr_Risco, _
'                PedLoc.Dsc_Compl_Risco, PedLoc.Val_Vr, PedLoc.NOSSO_NUMERO, PedLoc.Insp_Obr)
'            Next
'        End With
        Screen.MousePointer = vbDefault
    Else
        optCondicaoInspecao(e_StatusInspecao_ConfNormas).Value = False
        optCondicaoInspecao(e_StatusInspecao_Dispensada).Value = False
        optCondicaoInspecao(e_StatusInspecao_obrigatoria).Value = False
    End If
End Sub

Private Sub cmdSair_Click()
    Me.Visible = False
    frmT46V101A.SetFocus
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If gM46V999.gTipoAcesso <> DIGITACAO Then
        If UnloadMode = 0 And Toolbar.Enabled = True Then
            Cancel = 1
            MsgBox "Utilize o bot�o sair!", vbInformation, Me.Caption
            Exit Sub
        End If
    End If

End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button
        Case "Consultar"
            Call mfBuscaInspecao
        Case "Editar"
            If gM46V111.gCod_plano < 100 Then
                Call frmT46V208A.CarregarForm(mobjstcPed.ColecaoPedLoc, mColecaoInspecao, flxSolicInspecao.Row)
                Call Carregar_Form(mobjstcPed)
            End If
        Case "Sair"
            Salvar_Item
            Unload Me
    End Select
End Sub
