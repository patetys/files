VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmT46V144A 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Textos Pré-Definidos de Especificação"
   ClientHeight    =   6330
   ClientLeft      =   1620
   ClientTop       =   1575
   ClientWidth     =   8985
   Icon            =   "T46V144A.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   8985
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "Incluir"
      Height          =   285
      Left            =   90
      TabIndex        =   2
      Top             =   5985
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.TextBox TxtAtualEspec 
      Height          =   285
      Left            =   8460
      TabIndex        =   1
      Top             =   5940
      Visible         =   0   'False
      Width           =   465
   End
   Begin MSFlexGridLib.MSFlexGrid GrdEspecif 
      Height          =   3225
      Left            =   45
      TabIndex        =   0
      Tag             =   "0"
      Top             =   45
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   5689
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin Threed.SSCommand cmdOK 
      Height          =   300
      Left            =   4185
      TabIndex        =   3
      ToolTipText     =   "Volta à proposta."
      Top             =   5985
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&OK"
      ForeColor       =   8388608
   End
   Begin RichTextLib.RichTextBox txtDescricao 
      Height          =   2625
      Left            =   45
      TabIndex        =   4
      Top             =   3285
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   4630
      _Version        =   393217
      HideSelection   =   0   'False
      ScrollBars      =   3
      TextRTF         =   $"T46V144A.frx":030A
   End
End
Attribute VB_Name = "frmT46V144A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mpCarregarGrid()
    Dim lobjstc As stcA46V101B    'Objeto da classe de estrutura stcA46V101B.
    Dim lbdTab_Espec As ADODB.Connection    'Arquivo: P00Tarifa.
    Dim lMensagem As String     'Mensagem.
    Dim lrsTab_Espec As clsYasRecordSet    'Registro : Tbm_Espec_Padrao.
    Dim lLinha As Long    'Linhas.
    Dim lSelect As String

    If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTab_Espec, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If

    If frmT46V101A.cboPlano.ListIndex = -1 Then
        MsgBox "Nenhum Modalidade selecionada!"
        Exit Sub
    End If

    GrdEspecif.Rows = 2
    GrdEspecif.Row = 1
    Dim i As Integer
    For i = 0 To GrdEspecif.Cols - 1
        GrdEspecif.Col = i
        GrdEspecif.Text = ""
    Next

    lSelect = " Select Cod_Seq_Espec, Dsc_Item  from Tbm_Espec_Padrao "
    lSelect = lSelect & " Where cod_Ramo = " & gM46V999.gCod_Ramo
    lSelect = lSelect & " and Cod_Espec = " & TxtAtualEspec
    lSelect = lSelect & " and Cod_Modalidade = " & gM46V111.gCod_Plano
    lSelect = lSelect & " and Cod_Tarif = " & gM46V999.gCodProduto
    lSelect = lSelect & " order by Cod_Seq_Espec"

    If gM46V999.gfObtRegistro(lbdTab_Espec, lSelect, lrsTab_Espec, lMensagem) = False Then
        Call gM46V999.gpFecha2(lbdTab_Espec)
        Call gM46V999.gpGraLog(2, lMensagem)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    GrdEspecif.Rows = 1
    lLinha = 1
    txtDescricao.Visible = False
    Do While Not lrsTab_Espec.EOF
        Set lobjstc = New stcA46V101B
        GrdEspecif.Rows = lLinha + 1
        GrdEspecif.TextMatrix(lLinha, 0) = lrsTab_Espec("Cod_Seq_Espec")
        txtDescricao.TextRTF = lrsTab_Espec("Dsc_Item")
        GrdEspecif.TextMatrix(lLinha, 1) = Trim(txtDescricao.Text)
        GrdEspecif.TextMatrix(lLinha, 2) = lrsTab_Espec("Dsc_Item")
        lrsTab_Espec.MoveNext
        lLinha = lLinha + 1
    Loop
    Call gM46V999.gpFecha2(lbdTab_Espec)
    txtDescricao.Text = ""
    txtDescricao.Visible = True
    If GrdEspecif.Rows > 1 Then
        GrdEspecif.Row = 1
        GrdEspecif.Col = 0
        GrdEspecif_Click
    End If
End Sub

Private Sub CmdIncluir_Click()
    Dim lbdTab_Espec As ADODB.Connection    'Arquivo: P00Tarifa.
    Dim lMensagem As String     'Mensagem.
    Dim lrsTab_Espec As ADODB.Recordset    'Registro : Tbm_Espec_Padrao. 'Este recordSet esta usando a propriedade Filter (nao implementado no YasRecordSet)
    Dim lSelect As String

    If Trim(txtDescricao.Text) = "" Then Exit Sub

    lSelect = " Select *  from Tbm_Espec_Padrao "
    lSelect = lSelect & " Where cod_Ramo = " & gM46V999.gCod_Ramo
    lSelect = lSelect & " and Cod_Espec = " & TxtAtualEspec
    lSelect = lSelect & " and Cod_Modalidade = " & gM46V111.gCod_Plano
    lSelect = lSelect & " and Cod_Tarif = " & gM46V999.gCodProduto

    If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTab_Espec, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If

    Set lrsTab_Espec = New clsYasRecordSet
    lrsTab_Espec.CursorType = adOpenKeyset
    lrsTab_Espec.LockType = adLockOptimistic

    lrsTab_Espec.Open lSelect, lbdTab_Espec    'Este recordSet esta usando a propriedade Filter (nao implementado no YasRecordSet)

    lrsTab_Espec.AddNew
    lrsTab_Espec("Cod_Ramo") = gM46V999.gCod_Ramo
    lrsTab_Espec("Cod_Modalidade") = gM46V111.gCod_Plano
    lrsTab_Espec("Cod_Espec") = TxtAtualEspec
    lrsTab_Espec("Cod_Seq_Espec") = Int((99 * Rnd) + 1)
    lrsTab_Espec("Dsc_Item") = txtDescricao.TextRTF
    lrsTab_Espec("Cod_Tarif") = gM46V999.gCodProduto
    lrsTab_Espec.Update

    Call gM46V999.gpFecha2(lbdTab_Espec)
    mpCarregarGrid
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    mpCarregarGrid
End Sub

Private Sub Form_Load()
'1. Posiciona form.
    Call gM46V999.gpCenForm(Me)

    GrdEspecif.Cols = 3

    GrdEspecif.Row = 0
    GrdEspecif.Col = 0
    GrdEspecif.Text = "Item"
    GrdEspecif.Col = 1
    GrdEspecif.Text = "Descrição"

    GrdEspecif.ColWidth(0) = 700
    GrdEspecif.ColWidth(1) = 12300
    GrdEspecif.ColWidth(2) = 0
End Sub

Private Sub GrdEspecif_Click()
    txtDescricao.TextRTF = GrdEspecif.TextMatrix(GrdEspecif.RowSel, 2)
End Sub

Private Sub GrdEspecif_DblClick()
'frmT46V143A.RchFinal.TextRTF = GrdEspecif.TextMatrix(GrdEspecif.RowSel, 2)
    frmT46V153A.RchFinal.TextRTF = GrdEspecif.TextMatrix(GrdEspecif.RowSel, 2)
    Unload Me
End Sub

