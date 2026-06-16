VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT47V002A 
   Caption         =   "Selecione a classificação"
   ClientHeight    =   3690
   ClientLeft      =   1860
   ClientTop       =   1545
   ClientWidth     =   6585
   Icon            =   "T47V002.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3690
   ScaleWidth      =   6585
   Begin VB.ListBox lstClassificacao 
      ForeColor       =   &H00800000&
      Height          =   2790
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   6540
   End
   Begin VB.TextBox txtClassificacao 
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1200
      MaxLength       =   40
      TabIndex        =   0
      ToolTipText     =   "Informe os dados para a busca."
      Top             =   0
      Width           =   5340
   End
   Begin Threed.SSCommand cmdOk 
      Height          =   300
      Left            =   2760
      TabIndex        =   2
      ToolTipText     =   "Volta ao item com a classificação selecionada."
      Top             =   3240
      Width           =   1020
      _Version        =   65536
      _ExtentX        =   1799
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Ok"
      ForeColor       =   8388608
   End
   Begin VB.Label lblX 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Busca: "
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   0
      Left            =   0
      TabIndex        =   3
      ToolTipText     =   "Ordem da listagem da região."
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmT47V002A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ColecaoClassificacao    As Collection      'Coleção.
Private mDescricao              As String
Private mCodigo                 As String
Private mFinalizar              As Boolean
Private Sub cmdOK_Click()
Dim lobjstc                     As stcA46V101B      'Objeto da classe de estrutura stcA46V101A.

    If lstClassificacao.ListIndex = -1 Then
        If lstClassificacao.ListCount > 0 Then
            lstClassificacao.ListIndex = 0
        End If
    End If
    If lstClassificacao.ListIndex <> -1 Then
        Set lobjstc = ColecaoClassificacao.Item(lstClassificacao.List(lstClassificacao.ListIndex))
        mDescricao = lobjstc.Codigo & "-" & lobjstc.Descricao
        mCodigo = lobjstc.Chave
        'frmT47V001A.txtClassificacao.Text = lobjstc.Codigo & "-" & lobjstc.Descricao
        'frmT47V001A.txtClassificacao.Tag = lobjstc.Chave
    End If
    Unload Me
End Sub
Public Sub gpIntanciarForm(pCodClasf As String, pDescricao As String)
    mFinalizar = False
    Me.Show vbModal
    pDescricao = mDescricao
    pCodClasf = mCodigo
End Sub

Private Sub Form_Load()
    '1. Posiciona form.
    Call gM46V999.gpCenForm(Me)
    
    Call mpCarregarList
End Sub
Private Sub mpCarregarList()
Dim lobjstc         As stcA46V101B          'Objeto da classe de estrutura stcA46V101A.
Dim lbdTAB_Class    As ADODB.Connection     'Arquivo P00Tarifa.
Dim lMensagem       As String               'Mensagem.
Dim rsTab_Class     As clsYasRecordSet      'Registro, TAB_CLASF
Dim lCount          As Long                 'Contador.
Dim lSelect         As String

    
    Set ColecaoClassificacao = New Collection
    If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTAB_Class, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If
    lSelect = " Select distinct dsc_clasf, COD_CLASF, " & _
                "COD_COMPL_CLASF , tab_Clasf from Tbm_Classificacao " & _
                " Where cod_Ramo  = " & gM46V999.gCod_Ramo
    
    If frmT47V001A.cboPlano(0).ListIndex <> -1 Then
        Select Case frmT47V001A.cboPlano(0).ItemData(frmT47V001A.cboPlano(0).ListIndex)
            Case 1020, 1030, 1040, 1050, 1060
                lSelect = lSelect & " and  tab_clasf = " & frmT47V001A.cboPlano(0).ItemData(frmT47V001A.cboPlano(0).ListIndex)
            Case Else
                lSelect = lSelect & " and  tab_clasf = " & frmT47V001A.cboPlano(0).ItemData(frmT47V001A.cboPlano(0).ListIndex)
        End Select
    End If
    lSelect = lSelect & " and  cod_tarif  =  " & gM46V999.gCodProduto
    lSelect = lSelect & " and  cod_sequencia  =  " & gM46V999.gCodSequencia
    
    If gM46V999.gfObtRegistro(lbdTAB_Class, lSelect, rsTab_Class, lMensagem) = False Then
        Call gM46V999.gpFecha2(lbdTAB_Class)
        Call gM46V999.gpGraLog(2, lMensagem)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Do While Not rsTab_Class.EOF
        Set lobjstc = New stcA46V101B
        lobjstc.Descricao = Trim(rsTab_Class("DSC_CLASF"))
        lobjstc.Codigo = rsTab_Class("COD_CLASF") & "-" & Format(rsTab_Class("COD_COMPL_CLASF"), "00")
        lobjstc.Chave = rsTab_Class("COD_CLASF") & String(4 - Len(rsTab_Class("COD_CLASF")), " ") & Format(rsTab_Class("COD_COMPL_CLASF"), "00") & _
            Format(rsTab_Class("tab_Clasf"), "0000")
        ColecaoClassificacao.Add lobjstc, lobjstc.Codigo & " - " & lobjstc.Descricao
        rsTab_Class.MoveNext
    Loop
    lCount = 0
    For Each lobjstc In ColecaoClassificacao
        lCount = lCount + 1
        lstClassificacao.AddItem lobjstc.Codigo & " - " & lobjstc.Descricao
        lstClassificacao.ItemData(lstClassificacao.NewIndex) = lCount
        
        mCodigo = lobjstc.Chave
        mDescricao = lobjstc.Codigo & " - " & lobjstc.Descricao
    Next lobjstc
    
    If ColecaoClassificacao.Count = 1 Then
        mFinalizar = True
        'Unload Me
    End If
    
End Sub

Private Sub Form_Activate()
    
    If lstClassificacao.ListCount = 0 Then
        Call mpCarregarList
    End If
    
    If mFinalizar = True Then
        Unload Me
    End If
End Sub

Private Sub lstClassificacao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdOK_Click
    End If
End Sub

Private Sub txtClassificacao_KeyUp(KeyCode As Integer, Shift As Integer)
Dim lobjstc     As stcA46V101B          'Objeto da classe de estrutura stcA46V101A.

    lstClassificacao.Clear
    For Each lobjstc In ColecaoClassificacao
        If InStr(1, lobjstc.Codigo & " - " & lobjstc.Descricao, UCase(txtClassificacao.Text)) > 0 Then
            lstClassificacao.AddItem lobjstc.Codigo & " - " & lobjstc.Descricao
        End If
    Next lobjstc
End Sub

Private Sub lstClassificacao_DblClick()
    Call cmdOK_Click
End Sub


