VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmT47V022A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Texto Padrão para Especificação"
   ClientHeight    =   5715
   ClientLeft      =   1215
   ClientTop       =   1770
   ClientWidth     =   10110
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   10110
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid flxEspec 
      Height          =   375
      Left            =   3330
      TabIndex        =   10
      Top             =   5310
      Visible         =   0   'False
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   661
      _Version        =   393216
   End
   Begin Threed.SSCommand cmdExcluir 
      Height          =   375
      Left            =   8655
      TabIndex        =   9
      Top             =   45
      Width           =   1320
      _Version        =   65536
      _ExtentX        =   2328
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "Excluir Página"
   End
   Begin Threed.SSCommand cmdInserir 
      Height          =   375
      Left            =   7020
      TabIndex        =   8
      Top             =   45
      Width           =   1320
      _Version        =   65536
      _ExtentX        =   2328
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "Inserir Página"
   End
   Begin VB.TextBox txtPagina 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1035
      TabIndex        =   7
      Top             =   5310
      Width           =   555
   End
   Begin Threed.SSCommand cmdProx 
      Height          =   285
      Left            =   1575
      TabIndex        =   5
      Top             =   5310
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   503
      _StockProps     =   78
      Picture         =   "T47V022A.frx":0000
   End
   Begin Threed.SSCommand cmdAnt 
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Top             =   5310
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   503
      _StockProps     =   78
      Picture         =   "T47V022A.frx":00EB
   End
   Begin VB.TextBox txtCodDizer 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   45
      TabIndex        =   2
      Top             =   45
      Width           =   870
   End
   Begin Threed.SSCommand cmdFechar 
      Height          =   390
      Left            =   8775
      TabIndex        =   1
      ToolTipText     =   "Volta à proposta."
      Top             =   5265
      Width           =   1320
      _Version        =   65536
      _ExtentX        =   2328
      _ExtentY        =   688
      _StockProps     =   78
      Caption         =   "&Fechar"
      ForeColor       =   0
   End
   Begin Threed.SSCommand cmdBusca 
      Height          =   375
      Left            =   990
      TabIndex        =   3
      ToolTipText     =   "Volta à proposta."
      Top             =   45
      Width           =   1635
      _Version        =   65536
      _ExtentX        =   2884
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "&Busca Texto Padrão"
      ForeColor       =   8388608
   End
   Begin VB.TextBox txtTextoEspec 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   4800
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   450
      Width           =   10110
   End
   Begin VB.Label Label1 
      Caption         =   "Página:"
      Height          =   195
      Left            =   405
      TabIndex        =   6
      Top             =   5355
      Width           =   600
   End
End
Attribute VB_Name = "frmT47V022A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NumPag As Integer
Dim TotPag As Integer

Sub AtualizaListaEspec()
    flxEspec.TextMatrix(NumPag, 0) = NumPag
    flxEspec.TextMatrix(NumPag, 1) = gM46V999.CDblx(txtCodDizer)
    flxEspec.TextMatrix(NumPag, 2) = txtTextoEspec
End Sub

Sub MostraEspec()
     If TotPag > 0 Then
        txtCodDizer = flxEspec.TextMatrix(NumPag, 1)
        txtTextoEspec = flxEspec.TextMatrix(NumPag, 2)
        If gM46V999.CDblx(txtCodDizer) = 0 Then
            txtCodDizer.Enabled = True
            txtTextoEspec.Enabled = False
        Else
            txtCodDizer.Enabled = False
            txtTextoEspec.Enabled = True
        End If
     End If
End Sub

Private Sub cmdAnt_Click()
    If NumPag > 1 Then
        NumPag = NumPag - 1
        txtPagina = NumPag & "/" & TotPag
        MostraEspec
    End If
End Sub

Private Sub cmdBusca_Click()
Dim lbdTab_Ped_Dizer_Pad   As ADODB.Connection     'Arquivo: P00Tarifa.
Dim lrsTab_Ped_Dizer_Pad   As clsYasRecordSet      'Registro : Tab_Ped_Dizer_Pad.
Dim lMensagem           As String                  'Mensagem.
Dim lLinha              As Long                    'Linhas.
Dim lI                  As Long
Dim lClausula           As Integer
Dim lObjstcCorr         As stcA46V705B  'Objeto stcA46V705B  (corretor).
Dim lSelect     As String
   ' If Not txtTextoEspec.Enabled Then Exit Sub
    If Not txtCodDizer.Enabled Then Exit Sub
    
    txtCodDizer_LostFocus
    If txtCodDizer = "0" Then Exit Sub
    
    'Ficha 1000742 - Não exibir especificação Item 2 para Corretor Volkswagem
    If (gM46V999.CDblx(txtCodDizer) = 8695) Or (gM46V999.CDblx(txtCodDizer) = 8696) Then
        For Each lObjstcCorr In frmT47V001A.objstcPedido.ColecaoPedCorr
            If InStr(1, "0008001, 0008002, 0008003, 0008004, 0008005, 0008006, 0008008, 0008011, 0008014, 0008028", lObjstcCorr.Cod_Corr, vbTextCompare) > 0 Then
                MsgBox "Favor não utilizar esta especificação para Corretora Volkswagen!", vbInformation
                txtCodDizer.Text = ""
                Exit Sub
            End If
        Next
    End If

    gM46V999.gBaseLocal = False
    gM46V999.gTipP00MULTI = "1"
    If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTab_Ped_Dizer_Pad, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If
    

    lSelect = " Select * from Tab_Ped_Dizer_Pad  "
    lSelect = lSelect & " Where cod_Dizer_Pad = " & txtCodDizer
    lSelect = lSelect & " order by num_seq"
    
    
    If gM46V999.gfObtRegistro(lbdTab_Ped_Dizer_Pad, lSelect, lrsTab_Ped_Dizer_Pad, lMensagem) = False Then
        Call gM46V999.gpFecha2(lbdTab_Ped_Dizer_Pad)
        Call gM46V999.gpGraLog(2, lMensagem)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    gCarga = True
    txtTextoEspec = ""
    If lrsTab_Ped_Dizer_Pad.EOF Then
        MsgBox "Mensagem Padrão '" & txtCodDizer & "' não encontrada!", vbInformation
        txtCodDizer = "0"
    Else
        Do While Not lrsTab_Ped_Dizer_Pad.EOF
            If gM46V999.gfPreenchido(lrsTab_Ped_Dizer_Pad("Dsc_Dizer_Endo")) Then
                txtTextoEspec = txtTextoEspec & lrsTab_Ped_Dizer_Pad("Dsc_Dizer_Endo") & vbCrLf
            End If
            lrsTab_Ped_Dizer_Pad.MoveNext
        Loop
        txtTextoEspec = Mid(txtTextoEspec, 1, Len(txtTextoEspec) - 2)
        txtTextoEspec = RTrim(txtTextoEspec)
        txtCodDizer.Enabled = False
        txtTextoEspec.Enabled = True
        txtTextoEspec.SetFocus
    End If
    Call gM46V999.gpFecha2(lbdTab_Ped_Dizer_Pad)
    gCarga = False
    
    If txtCodDizer.Text = 8501 And gM46V999.gTipoAcesso = CALCULO Then
        If gM46V999.gCod_Ramo = e_CodRamo_Empresarial Then
            lClausula = 4
        ElseIf gM46V999.gCod_Ramo = e_CodRamo_Residencial Then
            lClausula = 1
        Else
            GoTo Fim
        End If
        If frmT47V023A.cboClausulas(1).ListIndex <> -1 Then
            If frmT47V023A.cboClausulas(1).ItemData(frmT47V023A.cboClausulas(1).ListIndex) = lClausula Then GoTo Fim
        Else
            For lI = 0 To frmT47V023A.cboClausulas(1).ListCount
                If frmT47V023A.cboClausulas(1).ItemData(lI) = lClausula Then
                    frmT47V023A.cboClausulas(1).ListIndex = lI
                    GoTo Fim
                End If
            Next lI
        End If
        If frmT47V023A.cboClausulas(2).ListIndex <> -1 Then
            If frmT47V023A.cboClausulas(2).ItemData(frmT47V023A.cboClausulas(2).ListIndex) = lClausula Then GoTo Fim
        Else
            For lI = 0 To frmT47V023A.cboClausulas(2).ListCount
                If frmT47V023A.cboClausulas(2).ItemData(lI) = lClausula Then
                    frmT47V023A.cboClausulas(2).ListIndex = lI
                    GoTo Fim
                End If
            Next lI
        End If
        If frmT47V023A.cboClausulas(3).ListIndex <> -1 Then
            If frmT47V023A.cboClausulas(3).ItemData(frmT47V023A.cboClausulas(3).ListIndex) = lClausula Then GoTo Fim
        Else
            For lI = 0 To frmT47V023A.cboClausulas(3).ListCount
                If frmT47V023A.cboClausulas(3).ItemData(lI) = lClausula Then
                    frmT47V023A.cboClausulas(3).ListIndex = lI
                    GoTo Fim
                End If
            Next lI
        End If
        If frmT47V023A.cboClausulas(4).ListIndex <> -1 Then
            If frmT47V023A.cboClausulas(4).ItemData(frmT47V023A.cboClausulas(4).ListIndex) = lClausula Then GoTo Fim
        Else
            For lI = 0 To frmT47V023A.cboClausulas(4).ListCount
                If frmT47V023A.cboClausulas(4).ItemData(lI) = lClausula Then
                    frmT47V023A.cboClausulas(4).ListIndex = lI
                    GoTo Fim
                End If
            Next lI
        End If
        If frmT47V023A.cboClausulas(5).ListIndex <> -1 Then
            If frmT47V023A.cboClausulas(5).ItemData(frmT47V023A.cboClausulas(5).ListIndex) = lClausula Then GoTo Fim
        Else
            For lI = 0 To frmT47V023A.cboClausulas(5).ListCount
                If frmT47V023A.cboClausulas(5).ItemData(lI) = lClausula Then
                    frmT47V023A.cboClausulas(5).ListIndex = lI
                    GoTo Fim
                End If
            Next lI
        End If
        If frmT47V023A.cboClausulas(6).ListIndex <> -1 Then
            If frmT47V023A.cboClausulas(6).ItemData(frmT47V023A.cboClausulas(6).ListIndex) = lClausula Then GoTo Fim
        Else
            For lI = 0 To frmT47V023A.cboClausulas(6).ListCount
                If frmT47V023A.cboClausulas(6).ItemData(lI) = lClausula Then
                    frmT47V023A.cboClausulas(6).ListIndex = lI
                    GoTo Fim
                End If
            Next lI
        End If
        If frmT47V023A.cboClausulas(7).ListIndex <> -1 Then
            If frmT47V023A.cboClausulas(7).ItemData(frmT47V023A.cboClausulas(7).ListIndex) = lClausula Then GoTo Fim
        Else
            For lI = 0 To frmT47V023A.cboClausulas(7).ListCount
                If frmT47V023A.cboClausulas(7).ItemData(lI) = lClausula Then
                    frmT47V023A.cboClausulas(7).ListIndex = lI
                    GoTo Fim
                End If
            Next lI
        End If
        If frmT47V023A.cboClausulas(8).ListIndex <> -1 Then
            If frmT47V023A.cboClausulas(8).ItemData(frmT47V023A.cboClausulas(8).ListIndex) = lClausula Then GoTo Fim
        Else
            For lI = 0 To frmT47V023A.cboClausulas(8).ListCount
                If frmT47V023A.cboClausulas(8).ItemData(lI) = lClausula Then
                    frmT47V023A.cboClausulas(8).ListIndex = lI
                    GoTo Fim
                End If
            Next lI
        End If
        If frmT47V023A.cboClausulas(9).ListIndex <> -1 Then
            If frmT47V023A.cboClausulas(9).ItemData(frmT47V023A.cboClausulas(9).ListIndex) = lClausula Then GoTo Fim
        Else
            For lI = 0 To frmT47V023A.cboClausulas(9).ListCount
                If frmT47V023A.cboClausulas(9).ItemData(lI) = lClausula Then
                    frmT47V023A.cboClausulas(9).ListIndex = lI
                    GoTo Fim
                End If
            Next lI
        End If
    End If
    frmT47V001A.barFerramentas.Buttons("Salvar").Enabled = True
    frmT47V001A.mnuArqSalvar.Enabled = True
Fim:
End Sub

Private Sub cmdExcluir_Click()
Dim i As Integer
    gCarga = True
    For i = 1 To flxEspec.Rows - 1
        If gM46V999.CDblx(txtCodDizer) = gM46V999.CDblx(flxEspec.TextMatrix(i, 1)) Then
           If flxEspec.Rows > 2 Then
                flxEspec.RemoveItem i
           Else
                flxEspec.TextMatrix(i, 0) = ""
                flxEspec.TextMatrix(i, 1) = ""
                flxEspec.TextMatrix(i, 2) = ""
           End If
           Exit For
        End If
    Next
    txtCodDizer = ""
    txtTextoEspec = ""
    NumPag = 1
    TotPag = flxEspec.Rows - 1
    txtPagina = NumPag & "/" & TotPag
    MostraEspec
    If txtTextoEspec.Enabled And txtTextoEspec.Visible Then
        txtTextoEspec.SetFocus
    End If
    gCarga = False
End Sub

Private Sub cmdFechar_Click()
    Me.Hide
End Sub


Private Sub cmdInserir_Click()
    gCarga = True
    If gM46V999.CDblx(txtCodDizer) = 0 Then Exit Sub
    AtualizaListaEspec
    flxEspec.AddItem "", flxEspec.Rows
    TotPag = flxEspec.Rows - 1
    NumPag = TotPag
    txtPagina = NumPag & "/" & TotPag
    MostraEspec
    txtCodDizer.SetFocus
    gCarga = False
End Sub


Private Sub cmdProx_Click()
    If NumPag < TotPag Then
        NumPag = NumPag + 1
        txtPagina = NumPag & "/" & TotPag
        MostraEspec
    End If
End Sub

Private Sub Form_Activate()
    If txtTextoEspec.Enabled And txtTextoEspec.Visible Then
        txtTextoEspec.SetFocus
    End If
    NumPag = 1
    TotPag = flxEspec.Rows - 1
    txtPagina = NumPag & "/" & TotPag
    MostraEspec
    If txtTextoEspec.Enabled And txtTextoEspec.Visible Then
        txtTextoEspec.SetFocus
    ElseIf txtCodDizer.Enabled And txtCodDizer.Visible Then
        txtCodDizer.SetFocus
    End If
    
Finalizar:
    If gM46V999.gDesabilitarForms Then
        If Me.Tag <> "Desabilitado" Then
            DesabilitaControles Me
            cmdFechar.Enabled = True
        End If
    Else
        If Me.Tag = "Desabilitado" Then
            HabilitaControles Me
        End If
    End If
    
End Sub

Private Sub Form_Load()
    Call gM46V999.gpCenForm(Me)
    flxEspec.FixedCols = 0
    flxEspec.FixedRows = 1
    flxEspec.Cols = 3
    flxEspec.Row = 1
End Sub





Private Sub txtCodDizer_LostFocus()
    txtCodDizer = gM46V999.CDblx(txtCodDizer)
End Sub


Private Sub txtTextoEspec_Change()
Dim AuxArray As Variant
Dim lNewTexto As String
Dim RestTexto As String
Dim lTextoOriginal As String
Dim i As Integer
Dim Pri As Boolean
Dim lTotLinhas As Integer
Dim PosCursor As Long
Static Alterando As Boolean


    If Alterando Or gCarga Then Exit Sub
    
    If Trim(txtTextoEspec) = "" Then Exit Sub
    
    Alterando = True
    lNewTexto = ""
    RestTexto = ""
    lTotLinhas = 0
    Pri = True
    

    PosCursor = txtTextoEspec.SelStart
    txtTextoEspec.Text = gM46V999.gfRetirarCaracterEspecial(txtTextoEspec.Text, True, False)
    lTextoOriginal = txtTextoEspec.Text
    
    If InStr(lTextoOriginal, vbCrLf) <> 0 Then
        AuxArray = Split(lTextoOriginal, vbCrLf)
        For i = 0 To UBound(AuxArray)
            AuxArray(i) = RestTexto & AuxArray(i)
            RestTexto = ""
            If Len(AuxArray(i)) > 80 Then
                RestTexto = Mid(AuxArray(i), 81)
                lNewTexto = lNewTexto & Mid(AuxArray(i), 1, 80) & vbCrLf
                If (Mid(lTextoOriginal, PosCursor + 1, 1) = vbCr Or Len(lTextoOriginal) = PosCursor) And Pri Then
                    PosCursor = PosCursor + 2
                    Pri = False
                End If
                lTotLinhas = lTotLinhas + 1
            Else
                lNewTexto = lNewTexto & AuxArray(i) & vbCrLf
                lTotLinhas = lTotLinhas + 1
            End If
        Next
        If RestTexto <> "" Then
            txtTextoEspec = lNewTexto & RestTexto
        Else
            txtTextoEspec = Mid(lNewTexto, 1, Len(lNewTexto) - 2)
        End If
    Else
        If Len(lTextoOriginal) > 80 Then
            Do While True
               lNewTexto = lNewTexto & Mid(lTextoOriginal, 1, 80) & vbCrLf
               lTextoOriginal = Mid(lTextoOriginal, 81)
               If Len(lTextoOriginal) = 0 Then
                    Exit Do
               End If
            Loop
            txtTextoEspec = Mid(lNewTexto, 1, Len(lNewTexto) - 2)
            PosCursor = PosCursor + 2
        End If
    End If
    If lTotLinhas > 20 Then
        MsgBox "Total de Linhas é superior ao limite máximo (20)!" & vbCr & "O excedente será desprezado!", vbInformation
    End If
    txtTextoEspec.SelStart = PosCursor
    Alterando = False
    Call gpLimPremios
End Sub

Private Sub txtTextoEspec_LostFocus()
    AtualizaListaEspec
End Sub


