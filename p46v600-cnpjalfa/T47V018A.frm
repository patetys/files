VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmT47V018A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relação de bens"
   ClientHeight    =   4080
   ClientLeft      =   1860
   ClientTop       =   1545
   ClientWidth     =   8535
   Icon            =   "T47V018A.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox CmbBens 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7785
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   3420
      Visible         =   0   'False
      Width           =   555
   End
   Begin MSFlexGridLib.MSFlexGrid flxBens 
      Height          =   3375
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "Para Inserir nova linha aperte Tecla INSERT"
      Top             =   0
      Width           =   8550
      _ExtentX        =   15081
      _ExtentY        =   5953
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin Threed.SSCommand cmdOk 
      Height          =   300
      Left            =   3788
      TabIndex        =   2
      ToolTipText     =   "Volta à proposta."
      Top             =   3720
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Ok"
      ForeColor       =   8388608
   End
   Begin Threed.SSCommand cmdNovo 
      Height          =   300
      Left            =   0
      TabIndex        =   3
      ToolTipText     =   "Novo bem."
      Top             =   3480
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Novo"
      ForeColor       =   8388608
   End
   Begin Threed.SSCommand cmdExcluir 
      Height          =   300
      Left            =   1080
      TabIndex        =   4
      ToolTipText     =   "Excluir o bem."
      Top             =   3480
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Excluir"
      ForeColor       =   8388608
   End
   Begin VB.CommandButton cmdFocus 
      Caption         =   "Command1"
      Height          =   495
      Left            =   5760
      TabIndex        =   5
      Top             =   2760
      Width           =   615
   End
End
Attribute VB_Name = "frmT47V018A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mLinhaFocus                 As Long             'Linha que recebeu o foco.
Dim mColunaFocus                As Long             'Coluna que recebeu o foco.
Dim lSortCriteria               As Long             'Utilizado para ordenação do grid
Dim lLastColSort                As Long             'Utilizado para ordenação do grid
Dim EstaCarregando              As Boolean

Private Sub CmbBens_KeyDown(KeyCode As Integer, Shift As Integer)
    With flxBens
        Select Case KeyCode
            Case 37
                .ColSel = .ColSel - 1
                .Col = .Col - 1
                If .Enabled And .Visible Then
                    .SetFocus
                End If
                DoEvents
                flxBens_Click
            Case 39
                .ColSel = .ColSel + 1
                .Col = .Col + 1
                flxBens_Click
            Case 13
                If .Row + 1 < .Rows Then
                    .RowSel = .RowSel + 1
                    .Row = .Row + 1
                    flxBens_Click
                End If
        End Select
    End With
End Sub


Private Sub CmbBens_LostFocus()
    If CmbBens.Visible Then
        If mLinhaFocus = 0 Then mLinhaFocus = flxBens.Row
        If CmbBens.ListIndex <> -1 Then
            flxBens.TextMatrix(mLinhaFocus, 1) = CmbBens.Text
            flxBens.TextMatrix(mLinhaFocus, 4) = CmbBens.ItemData(CmbBens.ListIndex)
        Else
            flxBens.TextMatrix(mLinhaFocus, 1) = ""
            flxBens.TextMatrix(mLinhaFocus, 4) = ""
        End If
        CmbBens.Visible = False
        CmbBens.ListIndex = -1
        If flxBens.Enabled And flxBens.Visible Then
            flxBens.SetFocus
        End If
        DoEvents
    End If
End Sub


Private Sub cmdExcluir_Click()
Dim lRowSel             As Integer      'Linha selecionada.
Dim lI                  As Integer      'Utilizada no for...loop.

    If Not flxBens.Enabled Then Exit Sub
    
    If flxBens.Rows > 2 Then
        lRowSel = flxBens.RowSel
        If flxBens.RowSel <> 0 Then
            flxBens.RemoveItem flxBens.RowSel
        End If
        For lI = lRowSel To flxBens.Rows - 1
            flxBens.TextMatrix(lI, 0) = Format(lI, "000")
        Next lI
        mLinhaFocus = mLinhaFocus - 1
    Else
        lI = 1
        flxBens.Row = lI
        flxBens.TextMatrix(lI, 0) = Format(lI, "000")
        flxBens.TextMatrix(lI, 1) = ""
        flxBens.TextMatrix(lI, 2) = ""
        flxBens.TextMatrix(lI, 3) = ""
        flxBens.TextMatrix(lI, 4) = ""
    End If
End Sub

Private Sub cmdFocus_GotFocus()
    With flxBens
        If Not .Visible Or Not .Enabled Then Exit Sub
        .SetFocus
        If .ColSel = 2 Then
            If .TextMatrix(.RowSel, 1) = "" And .TextMatrix(.RowSel, 2) = "" Then
                cmdOk.SetFocus
            Else
                .ColSel = 3
                .Col = 3
            End If
        Else
            If .TextMatrix(.RowSel, 1) = "" And .TextMatrix(.RowSel, 2) = "" Then
                cmdOk.SetFocus
            Else
                If .RowSel = .Rows - 1 Then
                    cmdNovo_Click
                End If
            End If
        End If
    End With
End Sub

Private Sub cmdNovo_Click()
    If Not flxBens.Enabled Then Exit Sub
    With flxBens
        If .Rows < 65 Then
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = Format(.Rows - 1, "000")
            .TextMatrix(.Rows - 1, 1) = ""
            .TextMatrix(.Rows - 1, 2) = ""
            If .Enabled And .Visible Then
                .SetFocus
            End If
            .Row = .Rows - 1
            .RowSel = .Rows - 1
            .ColSel = 1
            .Col = 1
            .RowHeight(.Row) = CmbBens.Height
            flxBens_Click
        End If
    End With
    
End Sub

Private Sub cmdOK_Click()
Dim lI          As Integer          'Utilizada no for...loop.
    'Verificar as informações digitadas.
    For lI = 1 To flxBens.Rows - 1
        If lI >= flxBens.Rows Then Exit For
        If flxBens.TextMatrix(lI, 1) = "" And flxBens.TextMatrix(lI, 2) = "" And gM46V999.CDblx(flxBens.TextMatrix(lI, 3)) = 0 Then
            If flxBens.Rows = 2 Then
                gCarga = True
                frmT47V001A.optBens(1).Value = True
                gCarga = False
                Exit For
            Else
                flxBens.RemoveItem lI
                lI = lI - 1
            End If
        ElseIf flxBens.TextMatrix(lI, 3) = "" Then
            flxBens.TextMatrix(lI, 3) = "0,00"
        End If
    Next lI
    Me.Hide
End Sub

Private Sub flxBens_Click()
    With flxBens
        If .ColSel = mColunaFocus And .RowSel = mLinhaFocus Then
            Exit Sub
        End If
        If EstaCarregando Then Exit Sub
        EstaCarregando = True
        If mLinhaFocus <> 0 Then
            If CmbBens.Visible Then
                CmbBens_LostFocus
            End If
            If mColunaFocus = 3 Then
                flxBens.TextMatrix(mLinhaFocus, 3) = Format(flxBens.TextMatrix(mLinhaFocus, 3), "#,###,##0.00")
            End If
        End If
    
        mColunaFocus = .ColSel
        mLinhaFocus = .RowSel
        
'## smrosa
        If .ColSel = 1 Then
            If gM46V999.gCod_Ramo = e_CodRamo_Roubos Then
                 'setar foco na coluna 2
            Else
                If CmbBens.ListCount <> 0 Then
                    CmbBens.Left = .CellLeft + 5
                    CmbBens.Width = .CellWidth + 10
                    CmbBens.Top = .Top + .CellTop + 5
                    CmbBens.ListIndex = gM46V999.gfRetornarListIndex(CmbBens, .TextMatrix(mLinhaFocus, 4))
                    CmbBens.ZOrder 0
                    CmbBens.Visible = True
                    If CmbBens.Enabled Then
                        CmbBens.SetFocus
                    End If
                Else
                    .ColSel = 2
                End If
            End If
        End If
        EstaCarregando = False
    End With

End Sub

Private Sub flxBens_KeyDown(KeyCode As Integer, Shift As Integer)
    With flxBens
        If .RowSel > 0 And .ColSel > 1 Then
            Select Case KeyCode
                Case 46
                    .Row = .RowSel
                    .Col = .ColSel
                    .Refresh
                    If .ColSel = 2 Then
                        .TextMatrix(.RowSel, .ColSel) = ""
                    Else
                        .TextMatrix(.RowSel, .ColSel) = "0,00"
                    End If
                Case 45
                   ' If .Row = .Rows - 1 Then
                        Call cmdNovo_Click
                   ' End If
            End Select
        End If
    End With
    Call gpHabSalvar
End Sub

Private Sub flxBens_KeyPress(KeyAscii As Integer)
    With flxBens
        If .RowSel > 0 And .ColSel > 1 Then
            Select Case KeyAscii
                Case 8, 44, 48 To 57
               ' If mLinhaFocus = .RowSel Then
                   ' mLinhaFocus = 0
                   ' mColunaFocus = 0
                   ' If .TextMatrix(.RowSel, .ColSel) = "" Then
                   '     Call cmdExcluir_Click
                   '     Exit Sub
                   ' End If
                   ' .TextMatrix(.RowSel, .ColSel) = ""
                   ' .Row = .RowSel
                   ' .Col = .ColSel
                   ' .Refresh
               ' Else
                    'If .ColSel = 2 Then
                        'If .TextMatrix(.RowSel, 2) = "" And .TextMatrix(.RowSel, 2) = "" Then
                            'Call cmdExcluir_Click
                            'Exit Sub
                     '   End If
                    'End If
               ' End If
            End Select
            Select Case KeyAscii
                Case 8              'BackSpace
                    If Len(.TextMatrix(.RowSel, .ColSel)) > 1 Then
                        .TextMatrix(.RowSel, .ColSel) = Left(.TextMatrix(.RowSel, .ColSel), Len(.TextMatrix(.RowSel, .ColSel)) - 1)
                    Else
                        .TextMatrix(.RowSel, .ColSel) = ""
                    End If
                Case 13             'Enter
                    If .RowSel < .Rows - 1 Then
                        .RowSel = .RowSel + 1
                        .Row = .RowSel
                    End If
                Case 44             'Vírgula
                    If InStr(1, .TextMatrix(.RowSel, .ColSel), Chr(44)) > 0 Then Exit Sub
                    .TextMatrix(.RowSel, .ColSel) = .TextMatrix(.RowSel, .ColSel) & Chr(KeyAscii)
                Case 48 To 57
                    If .ColSel = 2 Then
                        If Len(.TextMatrix(.RowSel, .ColSel)) < 60 Then
                            .TextMatrix(.RowSel, .ColSel) = .TextMatrix(.RowSel, .ColSel) & Chr(KeyAscii)
                        Else
                            Beep
                        End If
                    Else
                        If Len(.TextMatrix(.RowSel, .ColSel)) < 12 Then
                            .TextMatrix(.RowSel, .ColSel) = .TextMatrix(.RowSel, .ColSel) & Chr(KeyAscii)
                        Else
                            Beep
                        End If
                    End If
                Case 32, 33, 35 To 38, 40 To 47, 61, 65 To 93, 95, 97 To 123, 126
                    If .ColSel = 2 Then
                        If Len(.TextMatrix(.RowSel, .ColSel)) < 60 Then
                           .TextMatrix(.RowSel, .ColSel) = .TextMatrix(.RowSel, .ColSel) & Chr(KeyAscii)
                        Else
                            Beep
                        End If
                    End If
                Case 22
                    .TextMatrix(.RowSel, .ColSel) = .TextMatrix(.RowSel, .ColSel) & Clipboard.GetText
            End Select
            Call gpHabSalvar
        End If
    End With
    
End Sub

Private Sub flxBens_RowColChange()
    flxBens_Click
End Sub

Private Sub flxBens_SelChange()
    flxBens.ColWidth(4) = 0
End Sub

Private Sub Form_Load()
Dim BDP00Tarifa     As ADODB.Connection 'Arquivo - P00Tarifa.
Dim rsTbmBens       As clsYasRecordSet  'TBM_TIP_BENS: tabela.
Dim lMensagem       As String
Dim lCodSequencia   As String
Dim lSelect         As String
    Call gM46V999.gpCenForm(Me)
    With flxBens
        .FixedRows = 1
        .FixedCols = 1
        .Rows = 2
        .Cols = 5
        .TextMatrix(0, 0) = "Nº"
        .TextMatrix(0, 1) = "Tipo do Bem"
        
        '## smrosa
        If gM46V999.gCod_Ramo = e_CodRamo_Roubos Then
            .ColWidth(1) = 0
        Else
            .ColWidth(1) = 3000
        End If
        
        .TextMatrix(0, 2) = "Descrição do Bem"
        .TextMatrix(0, 3) = "Valor do bem"
        .TextMatrix(0, 4) = ""
        .ColWidth(0) = 660
        .ColWidth(2) = 6240
        .ColWidth(3) = 1515
        .ColWidth(4) = 0
        .TextMatrix(1, 0) = "001"
        .TextMatrix(1, 1) = ""
        .TextMatrix(1, 2) = ""
        .TextMatrix(1, 3) = ""
        .TextMatrix(1, 4) = ""
        .Row = 1
        .ColAlignment(2) = flexAlignLeftTop
    End With
    
    mLinhaFocus = 1
    
    'Abre arquivo.
    If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, BDP00Tarifa, lMensagem) = False Then
        Exit Sub
    End If
    
        '3. Obtém produto sem término de vigência ou com término de vigência superior à data do dia.
    lSelect = "SELECT * FROM Tbm_Tip_Bens "
    lSelect = lSelect & "  where Cod_Tarif = " & gM46V999.gCodProduto
    lSelect = lSelect & "    and Cod_Sequencia  <= " & gM46V999.gCodSequencia
    
    '## smrosa
    If gM46V999.gCodProduto >= 947 Then
        lSelect = lSelect & "    and Cod_Ramo  = " & gM46V999.gCod_Ramo
    End If
    
    lSelect = lSelect & "    order by Cod_Sequencia, Dsc_Bem desc "
              
    If gM46V999.gfObtRegistro(BDP00Tarifa, lSelect, rsTbmBens, lMensagem) = False Then
        Call gM46V999.gpFecha2(BDP00Tarifa)
        Exit Sub
    End If
    
    lCodSequencia = 0
    If Not rsTbmBens.EOF Then
        lCodSequencia = rsTbmBens("Cod_Sequencia")
    End If
    
    CmbBens.Clear
    Do While Not rsTbmBens.EOF
        If lCodSequencia = rsTbmBens("Cod_Sequencia") Then
            CmbBens.AddItem rsTbmBens("Dsc_Bem")
            CmbBens.ItemData(CmbBens.NewIndex) = rsTbmBens("Cod_Bem")
        End If
        rsTbmBens.MoveNext
    Loop
    
    Call gM46V999.gpFecha2(BDP00Tarifa)

    flxBens.RowHeight(1) = CmbBens.Height
    If CmbBens.ListCount = 0 Then
        flxBens.ColWidth(1) = 0
    End If
End Sub


