VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmT47V005A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Busca unidade de negócio"
   ClientHeight    =   4740
   ClientLeft      =   1860
   ClientTop       =   1545
   ClientWidth     =   5775
   Icon            =   "T47V005A.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSair 
      BackColor       =   &H000000FF&
      Cancel          =   -1  'True
      Height          =   300
      Left            =   0
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   300
   End
   Begin MSFlexGridLib.MSFlexGrid flxUniNegocio 
      Height          =   2870
      Left            =   100
      TabIndex        =   0
      Top             =   1800
      Width           =   5600
      _ExtentX        =   9869
      _ExtentY        =   5054
      _Version        =   393216
      Cols            =   4
      ForeColor       =   8388608
      ForeColorFixed  =   12582912
      ForeColorSel    =   12582912
      AllowUserResizing=   3
      Appearance      =   0
   End
   Begin Threed.SSPanel panX 
      Height          =   1200
      Index           =   2
      Left            =   3300
      TabIndex        =   1
      Top             =   500
      Width           =   2400
      _Version        =   65536
      _ExtentX        =   4233
      _ExtentY        =   2117
      _StockProps     =   15
      Caption         =   "Como"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      Alignment       =   6
      Begin VB.OptionButton optComo 
         Caption         =   "&Qualquer parte do nome"
         ForeColor       =   &H00800000&
         Height          =   200
         Index           =   2
         Left            =   100
         TabIndex        =   4
         ToolTipText     =   "Pesquisar qualquer parte do nome."
         Top             =   900
         Value           =   -1  'True
         Width           =   2200
      End
      Begin VB.OptionButton optComo 
         Caption         =   "&Início do nome"
         ForeColor       =   &H00800000&
         Height          =   200
         Index           =   1
         Left            =   100
         TabIndex        =   3
         ToolTipText     =   "Pesquisar início do nome."
         Top             =   600
         Width           =   2200
      End
      Begin VB.OptionButton optComo 
         Caption         =   "&Nome inteiro"
         ForeColor       =   &H00800000&
         Height          =   200
         Index           =   0
         Left            =   100
         TabIndex        =   2
         ToolTipText     =   "Pesquisar nome inteiro."
         Top             =   300
         Width           =   2200
      End
   End
   Begin Threed.SSPanel panX 
      Height          =   1200
      Index           =   0
      Left            =   100
      TabIndex        =   5
      Top             =   500
      Width           =   3100
      _Version        =   65536
      _ExtentX        =   5468
      _ExtentY        =   2117
      _StockProps     =   15
      Caption         =   "Procurar"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      Alignment       =   6
      Begin VB.TextBox txtNome 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   100
         MaxLength       =   30
         TabIndex        =   6
         ToolTipText     =   "Nome a ser pesquisado."
         Top             =   500
         Width           =   2900
      End
   End
   Begin MSComctlLib.ImageList imgLista 
      Left            =   5250
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T47V005A.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T47V005A.frx":0326
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T47V005A.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T47V005A.frx":0A92
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T47V005A.frx":0DAC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar barFerramentas 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imgLista"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Pesquisar"
            Object.ToolTipText     =   "Pesquisar unidades produtivas - F3."
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Selecionar"
            Object.ToolTipText     =   "Selecionar unidade produtiva - F6."
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Parar"
            Object.ToolTipText     =   "Interromper processo."
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair do programa."
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "frmT47V005A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mInterromper    As Boolean  'Indica que usuário clicou botão interromper.
Private mCod_Corretor   As Long     'Cod Corretor.
Private mCod_Unidade    As Integer  'Cod Unidade.
Private mCod_Produtor   As Integer  'Cod Produtor.

Public Sub mpInstanciarForm(pCod_Corretor As Long, ByRef pCod_Unidade As Integer, ByRef pCod_Produtor As Integer)
    'Posiciona as informações.
    mCod_Corretor = pCod_Corretor
    panX(0).Enabled = False
    panX(2).Enabled = False
    barFerramentas.Buttons(7).Enabled = False
    
    
    Me.Show vbModal, frmT47V004A
    
    'Retorna as informações.
    pCod_Unidade = mCod_Unidade
    pCod_Produtor = mCod_Produtor
End Sub

Private Sub barFerramentas_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case UCase(Button.Key)
        Case "PESQUISAR"
            Call mnuArqPesquisar_Click
        Case "SELECIONAR"
            Call mnuArqSelecionar_Click
        Case "PARAR"
            mInterromper = True
        Case "SAIR"
            Unload Me
    End Select
End Sub
Private Sub cmdSair_Click()
    'Observação: controle criado para usuário sair do FORM utilizando tecla ESC. Invisível pelo usuário.
    Unload Me
End Sub
Private Sub flxUniNegocio_DblClick()
    Call mnuArqSelecionar_Click
End Sub
Private Sub mpPesquisar()
    Dim lbdP0044700 As ADODB.Connection 'Arquivo - P0044700.
    Dim lMensagem   As String           'Mensagem.
    Dim lrsDEPT     As clsYasRecordSet  'Registro - YTAB_DEPT.
    Dim lSelect     As String
    Screen.MousePointer = vbHourglass

    '1. Seleciona registros na tabela e monta grade.
    If gM46V999.gfAbrBasDados("P0044700", gM46V999.gTipP0044700, gM46V999.gEndP0044700, lbdP0044700, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    lSelect = " SELECT YTAB_DEPT.NUM_DEPT, YTAB_DEPT.NOM_DEPT, "
    lSelect = lSelect & " YTAB_DEPT.COD_IDEN_DEPT, YTAB_UNID_CORR_PROD.COD_CORR, "
    lSelect = lSelect & " YTAB_UNID_CORR_PROD.COD_PROD, YTAB_UNID_CORR_PROD.COD_UNID"
    lSelect = lSelect & " FROM YTAB_UNID_CORR_PROD INNER JOIN YTAB_DEPT ON "
    lSelect = lSelect & " YTAB_UNID_CORR_PROD.COD_UNID = YTAB_DEPT.NUM_DEPT"
    lSelect = lSelect & " WHERE YTAB_DEPT.COD_IDEN_DEPT In (4,10) AND "
    lSelect = lSelect & " YTAB_UNID_CORR_PROD.COD_CORR=" & mCod_Corretor
    lSelect = lSelect & " ORDER BY YTAB_DEPT.NOM_DEPT;"

    With flxUniNegocio
        .Rows = 1
        If gM46V999.gfObtRegistro(lbdP0044700, lSelect, lrsDEPT, lMensagem) = False Then
            Call gM46V999.gpFecha2(lbdP0044700)
            Call gM46V999.gpGraLog(2, lMensagem)
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        If lrsDEPT.EOF = True Then
            lMensagem = "FNLXXXX - " & gM46V999.gPrefixo & _
                        "Unidade de negócios não localizada"
            If optComo(0) = True Then       'Nome inteiro.
                lMensagem = lMensagem & " (nome inteiro = " & Trim$(txtNome.Text) & ")."
            Else
                If optComo(1) = True Then   'Início do nome.
                    lMensagem = lMensagem & " (início do nome = " & Trim$(txtNome.Text) & ")."
                Else
                    If optComo(2) = True Then   'Qualquer parte do nome.
                        lMensagem = lMensagem & " (qualquer parte do nome = " & Trim$(txtNome.Text) & ")."
                    End If
                End If
            End If
            Call gM46V999.gpGraLog(1, lMensagem)
        Else
            Do While Not lrsDEPT.EOF
                .Rows = .Rows + 1
                .Row = .Rows - 1
                .Col = 0
                .Text = .Row
                .Col = 1
                .Text = lrsDEPT!NUM_DEPT
                .Col = 2
                .Text = " " & lrsDEPT!NOM_DEPT
                .Col = 3
                .Text = lrsDEPT!Cod_Prod
                If .Rows > 11 Then
                    .TopRow = .Rows - 10
                End If
                DoEvents
                If mInterromper = True Then
                    mInterromper = False
                    Exit Do
                End If
                lrsDEPT.MoveNext
            Loop
            .TopRow = 1
        End If
        Call gM46V999.gpFecha2(lbdP0044700)

        '3. Reposiciona janela.
        If .Rows <> 1 Then
            
            barFerramentas.Buttons(2).Enabled = False   'Pesquisar.
            barFerramentas.Buttons(4).Enabled = True    'Selecionar.
        End If
        .Refresh
    End With

    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    Screen.MousePointer = vbHourglass

    '1. Centraliza form.
    Call gM46V999.gpCenForm(Me)

    '2. Prepara grade.
    With flxUniNegocio
        .Row = 0
        'a) Posiciona largura das colunas.
        .ColWidth(0) = 500
        .ColWidth(1) = 0
        .ColWidth(2) = 4800
        .ColWidth(3) = 0
        'b) Posiciona títulos.
        .Col = 1
        .Text = "Código da unidade de negócios"
        .Col = 2
        .Text = "Unidade de Negócios"
        .Col = 3
        .Text = "Produtor"
    End With

    '3. Desabilita opção de seleção.
    
    barFerramentas.Buttons(4).Enabled = False   'Selecionar.
    mInterromper = False
    If mCod_Corretor = 0 Then
        Call mnuArqPesquisar_Click
    Else
        Call mpPesquisar
    End If

    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mCod_Corretor = 0
End Sub

Private Sub mnuArqPesquisar_Click()
    Dim lbdP0044700 As ADODB.Connection 'Arquivo - P0044700.
    Dim lMensagem   As String           'Mensagem.
    Dim lrsDEPT     As clsYasRecordSet  'Registro - YTAB_DEPT.
    Dim lSelect     As String
    Screen.MousePointer = vbHourglass

    '1. Seleciona registros na tabela e monta grade.
    If gM46V999.gfAbrBasDados("P0044700", gM46V999.gTipP0044700, gM46V999.gEndP0044700, lbdP0044700, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    lSelect = "SELECT NUM_DEPT, NOM_DEPT, COD_IDEN_DEPT FROM YTAB_DEPT WHERE COD_IDEN_DEPT IN (4, 10)"
    If gM46V999.gfPreenchido(txtNome.Text) Then
        lSelect = lSelect & " AND NOM_DEPT "
        If optComo(0) = True Then       'Nome inteiro.
            lSelect = lSelect & " = '" & gM46V999.gfForAspas(txtNome.Text) & "'"
        ElseIf optComo(1) = True Then   'Início do nome.
            lSelect = lSelect & "LIKE '" & gM46V999.gfForAspas(txtNome.Text) & "%'"
        ElseIf optComo(2) = True Then   'Qualquer parte do nome.
            lSelect = lSelect & "LIKE '%" & gM46V999.gfForAspas(txtNome.Text) & "%'"
        End If
    End If
    lSelect = lSelect & " ORDER BY NOM_DEPT"
    With flxUniNegocio
        .Rows = 1
        If gM46V999.gfObtRegistro(lbdP0044700, lSelect, lrsDEPT, lMensagem) = False Then
            Call gM46V999.gpFecha2(lbdP0044700)
            Call gM46V999.gpGraLog(2, lMensagem)
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        If lrsDEPT.EOF = True Then
            lMensagem = "FNLXXXX - " & gM46V999.gPrefixo & _
                        "Unidade de negócios não localizada"
            If optComo(0) = True Then       'Nome inteiro.
                lMensagem = lMensagem & " (nome inteiro = " & Trim$(txtNome.Text) & ")."
            Else
                If optComo(1) = True Then   'Início do nome.
                    lMensagem = lMensagem & " (início do nome = " & Trim$(txtNome.Text) & ")."
                Else
                    If optComo(2) = True Then   'Qualquer parte do nome.
                        lMensagem = lMensagem & " (qualquer parte do nome = " & Trim$(txtNome.Text) & ")."
                    End If
                End If
            End If
            Call gM46V999.gpGraLog(1, lMensagem)
        Else
            Do While Not lrsDEPT.EOF
                .Rows = .Rows + 1
                .Row = .Rows - 1
                .Col = 0
                .Text = .Row
                .Col = 1
                .Text = lrsDEPT!NUM_DEPT
                .Col = 2
                .Text = " " & lrsDEPT!NOM_DEPT
                If .Rows > 11 Then
                    .TopRow = .Rows - 10
                End If
                DoEvents
                If mInterromper = True Then
                    mInterromper = False
                    Exit Do
                End If
                lrsDEPT.MoveNext
            Loop
            .TopRow = 1
        End If
        Call gM46V999.gpFecha2(lbdP0044700)

        '3. Reposiciona janela.
        If .Rows <> 1 Then
            barFerramentas.Buttons(2).Enabled = False   'Pesquisar.
            barFerramentas.Buttons(4).Enabled = True    'Selecionar.
        End If
    End With

    Screen.MousePointer = vbDefault
End Sub
Private Sub mnuArqSelecionar_Click()
    Screen.MousePointer = vbHourglass

    With flxUniNegocio
        '1. Se grade vazia, encerra processo.
        If .Rows = 1 Then
            Unload Me
            Exit Sub
        End If

        '2. Se indicada linha de título, encerra processo.
        If .Row = 0 Then
            Unload Me
            Exit Sub
        End If

        '3. Posiciona unidade de negócios.
        .Col = 1        'Código da unidade de negócios.
        frmT47V004A.lblNomUnidade.Tag = .Text
        mCod_Unidade = .Text
        .Col = 2
        frmT47V004A.lblNomUnidade.Caption = .Text
        .Col = 3
        mCod_Produtor = Val(.Text)
    End With

    Unload Me
    Screen.MousePointer = vbDefault
End Sub
Private Sub mnuArqSair_Click()
    Unload Me
End Sub
Private Sub optComo_Click(Index As Integer)
    flxUniNegocio.Rows = 1
    
    barFerramentas.Buttons(2).Enabled = True    'Pesquisar.
    barFerramentas.Buttons(4).Enabled = False   'Selecionar.
End Sub
Private Sub optComo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub
Private Sub txtNome_Change()
    flxUniNegocio.Rows = 1
    
    barFerramentas.Buttons(2).Enabled = True    'Pesquisar.
    barFerramentas.Buttons(4).Enabled = False   'Selecionar.
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

Private Sub txtNome_LostFocus()
    txtNome.Text = gM46V999.gfRetirarCaracterEspecial(txtNome.Text, True)
End Sub


