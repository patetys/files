VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmT46V180A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Prolabore"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8190
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
   Begin Threed.SSCommand cmdOk 
      Height          =   300
      Left            =   3720
      TabIndex        =   1
      Top             =   1965
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1764
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Ok"
      ForeColor       =   8388608
   End
   Begin VB.ComboBox cmbCorretores 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.ComboBox cmbProlaboristas 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4425
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox txtPercentual 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   3960
      MaxLength       =   60
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid flxProlaboristas 
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Para Inserir nova linha aperte Tecla INSERT, Para Excluir aperte Tecla DELETE"
      Top             =   0
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   3413
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin Threed.SSCommand cmdAdicionar 
      Height          =   300
      Left            =   600
      TabIndex        =   2
      Top             =   1965
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1764
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Inserir"
      ForeColor       =   8388608
   End
   Begin Threed.SSCommand cmdExcluir 
      Height          =   300
      Left            =   2160
      TabIndex        =   3
      Top             =   1965
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1764
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Excluir"
      ForeColor       =   8388608
   End
End
Attribute VB_Name = "frmT46V180A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mLinhaFocus As Long    'Linha que recebeu o foco.
Dim mColunaFocus As Long    'Coluna que recebeu o foco.
Dim lSortCriteria As Long    'Utilizado para ordenação do grid
Dim lLastColSort As Long    'Utilizado para ordenação do grid
Dim EstaCarregando As Boolean
Dim LimiteMaximoGeral As String
Dim LimiteMaximoEspecifico As String

Private Sub cmbCorretores_Click()
    If cmbCorretores.ListIndex <> -1 And flxProlaboristas.TextMatrix(mLinhaFocus, 2) <> "" Then
        If cmbCorretores.ItemData(cmbCorretores.ListIndex) <> flxProlaboristas.TextMatrix(mLinhaFocus, 2) Then
            flxProlaboristas.TextMatrix(mLinhaFocus, 3) = ""
            flxProlaboristas.TextMatrix(mLinhaFocus, 4) = ""
            flxProlaboristas.TextMatrix(mLinhaFocus, 5) = ""
        End If
    End If
End Sub

Private Sub cmbCorretores_DropDown()
    cmbCorretores.Clear
    Dim i As Integer
    For i = 1 To frmT46V102A.txtCorretor.Count
        If frmT46V102A.txtCorretor(i).Tag <> "" Then
            cmbCorretores.AddItem frmT46V102A.txtCorretor(i).Text
            cmbCorretores.ItemData(cmbCorretores.NewIndex) = frmT46V102A.txtCorretor(i).Tag
        End If
    Next i

    If (flxProlaboristas.TextMatrix(mLinhaFocus, 2)) <> "" Then
        cmbCorretores.ListIndex = gM46V999.gfRetornarListIndex(cmbCorretores, flxProlaboristas.TextMatrix(mLinhaFocus, 2))
    End If

End Sub

Private Sub cmbCorretores_KeyDown(KeyCode As Integer, Shift As Integer)
    With flxProlaboristas
        Select Case KeyCode
        Case 37
            .ColSel = .ColSel - 1
            .Col = .Col - 1
            If .Enabled And .Visible Then
                .SetFocus
            End If
            DoEvents
            flxProlaboristas_Click
        Case 39
            .ColSel = .ColSel + 1
            .Col = .Col + 1
            flxProlaboristas_Click
        Case 13
            If .Row + 1 < .Rows Then
                .RowSel = .RowSel + 1
                .Row = .Row + 1
                flxProlaboristas_Click
            End If
        End Select
    End With
End Sub

Private Sub cmbCorretores_LostFocus()
    If cmbCorretores.Visible Then
        If mLinhaFocus = 0 Then mLinhaFocus = flxProlaboristas.Row
        If cmbCorretores.ListIndex <> -1 Then
            flxProlaboristas.TextMatrix(mLinhaFocus, 1) = cmbCorretores.Text
            flxProlaboristas.TextMatrix(mLinhaFocus, 2) = cmbCorretores.ItemData(cmbCorretores.ListIndex)
        End If
        cmbCorretores.Visible = False
        cmbCorretores.ListIndex = -1
        If flxProlaboristas.Enabled And flxProlaboristas.Visible Then
            flxProlaboristas.SetFocus
        End If
        DoEvents
    End If
End Sub

Private Sub cmbProlaboristas_DropDown()
    Dim lMensagem As String    'Mensagem
    Dim lI As Double             'Utilizado com For...Next.
    Dim lSaida() As Variant    'Resultado da busca de corretores.
    Dim lQtdeRegistros As Double    'Quantidade de corretores selecionados.

    cmbProlaboristas.Clear

    If gM46V999.gfObtBuscaParceiroProlaborista(2, flxProlaboristas.TextMatrix(mLinhaFocus, 2), lSaida, lMensagem) = False Then
        Call gM46V999.gpGraLog(1, lMensagem)
        flxProlaboristas.TextMatrix(mLinhaFocus, 3) = ""
        flxProlaboristas.TextMatrix(mLinhaFocus, 4) = ""
        cmbProlaboristas.Clear
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    lQtdeRegistros = UBound(lSaida)

    cmbProlaboristas.AddItem "NÃO DEFINIDO"
    cmbProlaboristas.ItemData(cmbProlaboristas.NewIndex) = 0

    For lI = 1 To lQtdeRegistros
        cmbProlaboristas.AddItem lSaida(lI, 2)
        cmbProlaboristas.ItemData(cmbProlaboristas.NewIndex) = lSaida(lI, 1)
    Next lI

    If (flxProlaboristas.TextMatrix(mLinhaFocus, 4)) <> "" Then
        cmbProlaboristas.ListIndex = gM46V999.gfRetornarListIndex(cmbProlaboristas, flxProlaboristas.TextMatrix(mLinhaFocus, 4))
    End If

End Sub

Private Sub cmbProlaboristas_KeyDown(KeyCode As Integer, Shift As Integer)
    With flxProlaboristas
        Select Case KeyCode
        Case 37
            .ColSel = .ColSel - 1
            .Col = .Col - 1
            If .Enabled And .Visible Then
                .SetFocus
            End If
            DoEvents
            flxProlaboristas_Click
        Case 39
            .ColSel = .ColSel + 1
            .Col = .Col + 1
            flxProlaboristas_Click
        Case 13
            If .Row + 1 < .Rows Then
                .RowSel = .RowSel + 1
                .Row = .Row + 1
                flxProlaboristas_Click
            End If
        End Select
    End With
End Sub

Private Sub cmbProlaboristas_LostFocus()
    If cmbProlaboristas.Visible Then
        If mLinhaFocus = 0 Then mLinhaFocus = flxProlaboristas.Row
        If cmbProlaboristas.ListIndex <> -1 Then
            flxProlaboristas.TextMatrix(mLinhaFocus, 3) = cmbProlaboristas.Text
            flxProlaboristas.TextMatrix(mLinhaFocus, 4) = cmbProlaboristas.ItemData(cmbProlaboristas.ListIndex)
        End If
        cmbProlaboristas.Visible = False
        cmbProlaboristas.ListIndex = -1
        If flxProlaboristas.Enabled And flxProlaboristas.Visible Then
            flxProlaboristas.SetFocus
        End If
        DoEvents
    End If
End Sub

Private Sub cmdAdicionar_Click()
    If Not flxProlaboristas.Enabled Then Exit Sub
    With flxProlaboristas
        If .Rows < 65 Then
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = Format(.Rows - 1, "000")
            .TextMatrix(.Rows - 1, 1) = ""
            .TextMatrix(.Rows - 1, 2) = ""
            .TextMatrix(.Rows - 1, 3) = ""
            .TextMatrix(.Rows - 1, 4) = ""
            .TextMatrix(.Rows - 1, 5) = ""
            If .Enabled And .Visible Then
                .SetFocus
            End If
            .Row = .Rows - 1
            .RowSel = .Rows - 1
            .ColSel = 1
            .Col = 1
            .RowHeight(.Row) = cmbCorretores.Height
            flxProlaboristas_Click
        End If
    End With
End Sub

Private Sub cmdCancelar_Click()
    frmT46V102A.optProlabore(1).Value = True
    Me.Hide
End Sub

Private Sub cmdExcluir_Click()
    Dim lRowSel As Integer    'Linha selecionada.
    Dim lI As Integer                 'Utilizada no for...loop.

    If Not flxProlaboristas.Enabled Then Exit Sub

    If flxProlaboristas.Rows > 2 Then
        lRowSel = flxProlaboristas.RowSel
        If flxProlaboristas.RowSel <> 0 Then
            flxProlaboristas.RemoveItem flxProlaboristas.RowSel
        End If
        For lI = lRowSel To flxProlaboristas.Rows - 1
            flxProlaboristas.TextMatrix(lI, 0) = Format(lI, "000")
        Next lI
        mLinhaFocus = mLinhaFocus - 1
        cmbCorretores.Visible = False
    Else
        lI = 1
        flxProlaboristas.Row = lI
        flxProlaboristas.TextMatrix(lI, 0) = Format(lI, "000")
        flxProlaboristas.TextMatrix(lI, 1) = ""
        flxProlaboristas.TextMatrix(lI, 2) = ""
        flxProlaboristas.TextMatrix(lI, 3) = ""
        flxProlaboristas.TextMatrix(lI, 4) = ""
        flxProlaboristas.TextMatrix(lI, 5) = ""
    End If
End Sub

Private Sub cmdNovo_Click()
    If Not flxProlaboristas.Enabled Then Exit Sub
    With flxProlaboristas
        If .Rows < 65 Then
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = Format(.Rows - 1, "000")
            .TextMatrix(.Rows - 1, 1) = ""
            .TextMatrix(.Rows - 1, 2) = ""
            .TextMatrix(.Rows - 1, 3) = ""
            .TextMatrix(.Rows - 1, 4) = ""
            .TextMatrix(.Rows - 1, 5) = ""
            If .Enabled And .Visible Then
                .SetFocus
            End If
            .Row = .Rows - 1
            .RowSel = .Rows - 1
            .ColSel = 1
            .Col = 1
            .RowHeight(.Row) = cmbProlaboristas.Height
            flxProlaboristas_Click
        End If
    End With

End Sub

Private Sub cmdOK_Click()
    Dim lI As Integer
    Dim blnFechar As Boolean
    blnFechar = False

    For lI = 1 To flxProlaboristas.Rows - 1

        If lI >= flxProlaboristas.Rows Then Exit For

        If flxProlaboristas.Rows = 2 And flxProlaboristas.TextMatrix(lI, 2) = "" And flxProlaboristas.TextMatrix(lI, 4) = "" And gM46V999.CDblx(flxProlaboristas.TextMatrix(lI, 5)) = 0 Then
            gCarga = True
            frmT46V102A.optProlabore(1).Value = True
            gCarga = False
            blnFechar = True
            Exit For
        Else
            gCarga = True
            frmT46V102A.optProlabore(0).Value = True
            gCarga = False
            Exit For
        End If

    Next lI

    Call gpHabSalvar

    If (blnFechar) Then
        Unload Me
    Else
        Me.Hide
    End If
End Sub

Private Sub flxProlaboristas_Click()
    Dim lTexto As String

    lTexto = txtPercentual.Text
    If Trim(lTexto) <> "" Then
        flxProlaboristas.TextMatrix(mLinhaFocus, mColunaFocus) = lTexto
    End If

    With flxProlaboristas
        If EstaCarregando Then Exit Sub
        EstaCarregando = True
        If mLinhaFocus <> 0 Then

            If cmbCorretores.Visible Then
                cmbCorretores_LostFocus
            ElseIf cmbProlaboristas.Visible Then
                cmbProlaboristas_LostFocus
            ElseIf mColunaFocus = 5 Then
                flxProlaboristas.TextMatrix(mLinhaFocus, 5) = Format(flxProlaboristas.TextMatrix(mLinhaFocus, 5), "#,###,##0.00")
            End If

        End If

        mColunaFocus = .ColSel
        mLinhaFocus = .RowSel

        Select Case .ColSel

        Case 1
            cmbCorretores.Left = .CellLeft
            cmbCorretores.Width = .CellWidth
            cmbCorretores.Top = .Top + .CellTop
            cmbCorretores.ListIndex = gM46V999.gfRetornarListIndex(cmbCorretores, .TextMatrix(mLinhaFocus, 2))
            cmbCorretores.ZOrder 0
            cmbCorretores.Visible = True
            If cmbCorretores.Enabled Then
                cmbCorretores.SetFocus
            End If
        Case 3
            cmbProlaboristas.Left = .CellLeft
            cmbProlaboristas.Width = .CellWidth
            cmbProlaboristas.Top = .Top + .CellTop
            cmbProlaboristas.ListIndex = gM46V999.gfRetornarListIndex(cmbProlaboristas, .TextMatrix(mLinhaFocus, 4))
            cmbProlaboristas.ZOrder 0
            cmbProlaboristas.Visible = True
            If cmbProlaboristas.Enabled Then
                cmbProlaboristas.SetFocus
            End If
        End Select

        EstaCarregando = False

        .ColAlignment(2) = flexAlignLeftTop

    End With

End Sub

Private Sub flxProlaboristas_DblClick()
    mLinhaFocus = flxProlaboristas.Row
    mColunaFocus = flxProlaboristas.Col
    txtPercentual.Visible = False

    ExibirCelula
    With txtPercentual
        If .Visible Then
            .Text = flxProlaboristas.TextMatrix(mLinhaFocus, mColunaFocus)
            .SelStart = Len(.Text) + 1
        End If
    End With
End Sub

Private Sub flxProlaboristas_KeyDown(KeyCode As Integer, Shift As Integer)
    With flxProlaboristas
        If .RowSel > 0 And .ColSel > 1 Then
            Select Case KeyCode
            Case 46
                .Row = .RowSel
                .Col = .ColSel
                .Refresh
                If .ColSel = 5 Then
                    .TextMatrix(.RowSel, .ColSel) = ""
                Else
                    .TextMatrix(.RowSel, .ColSel) = "0,00"
                End If
            Case 45
                Call cmdNovo_Click
            End Select
        End If
    End With
    Call gpHabSalvar
End Sub

Private Sub flxProlaboristas_KeyPress(KeyAscii As Integer)
    With flxProlaboristas
        If .RowSel > 0 And .ColSel > 1 Then
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

            Case 22
                .TextMatrix(.RowSel, .ColSel) = .TextMatrix(.RowSel, .ColSel) & Clipboard.GetText

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
            End Select
            Call gpHabSalvar
        End If
    End With

End Sub

Private Sub flxProlaboristas_RowColChange()
    flxProlaboristas_Click
End Sub

Private Sub Form_Activate()
    frmT46V100A.panForm.Caption = Mid$(Me.Name, 4)
Finalizar:
    If gM46V999.gDesabilitarForms Or gM46V999.gDesabFormsCotacao Then
        If Me.Tag <> "Desabilitado" Then
            DesabilitaControles Me
            cmdOk.Enabled = True
        End If
    Else
        If Me.Tag = "Desabilitado" Then
            HabilitaControles Me
        End If
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyInsert Then cmdAdicionar_Click
    If KeyCode = vbKeyDelete Then cmdExcluir_Click

End Sub

Private Sub Form_Load()
    Dim lMensagem As String    'Mensagem
    Dim lI As Double             'Utilizado com For...Next.
    Dim lSaida() As Variant    'Resultado da busca de corretores.
    Dim lQtdeRegistros As Double    'Quantidade de corretores selecionados.

    Call gM46V999.gpCenForm(Me)
    With flxProlaboristas
        .FixedRows = 1
        .FixedCols = 1
        .Rows = 2
        .Cols = 6
        .TextMatrix(0, 0) = "Nº"
        .TextMatrix(0, 1) = "Corretor"
        .TextMatrix(0, 2) = ""
        .TextMatrix(0, 3) = "Prolabore"
        .TextMatrix(0, 4) = "Prolabore"
        .TextMatrix(0, 5) = "Percentual"
        .ColWidth(0) = 400
        .ColWidth(1) = 3000
        .ColWidth(2) = 0
        .ColWidth(3) = 3645
        .ColWidth(4) = 0
        .ColWidth(5) = 1000
        .TextMatrix(1, 0) = "001"
        .TextMatrix(1, 1) = ""
        .TextMatrix(1, 2) = ""
        .TextMatrix(1, 3) = ""
        .TextMatrix(1, 4) = ""
        .TextMatrix(1, 5) = ""
        .Row = 1
        .ColAlignment(2) = flexAlignLeftTop
    End With

    txtPercentual.Visible = False
    mLinhaFocus = 1

    cmbProlaboristas.Clear

    cmbCorretores.Clear

    flxProlaboristas.RowHeight(1) = cmbCorretores.Height

End Sub

Private Sub ExibirCelula()
    Dim lOK As Boolean
    If flxProlaboristas.Col <= flxProlaboristas.FixedCols - 1 Or flxProlaboristas.Row <= flxProlaboristas.FixedRows - 1 Then
        Exit Sub
    End If

    If lOK Then Exit Sub

    lOK = True
    txtPercentual.Visible = False

    mLinhaFocus = flxProlaboristas.Row
    mColunaFocus = flxProlaboristas.Col

    Select Case mLinhaFocus

    Case Else
        txtPercentual.Move flxProlaboristas.CellLeft + flxProlaboristas.Left, flxProlaboristas.CellTop + flxProlaboristas.Top, flxProlaboristas.CellWidth, flxProlaboristas.CellHeight
        txtPercentual.Text = flxProlaboristas.Text

        If Len(flxProlaboristas.Text) = 0 Then
            If mLinhaFocus > 1 Then
                txtPercentual.Text = flxProlaboristas.TextMatrix(mLinhaFocus - 1, mColunaFocus)
            End If
        End If

        txtPercentual.Visible = True

        If txtPercentual.Visible Then
            txtPercentual.ZOrder
            txtPercentual.SetFocus
        End If
    End Select

    lOK = False

End Sub

Private Sub txtPercentual_KeyPress(KeyAscii As Integer)
    Dim strMensagem As String
    ' ao pressionar ENTER aceitar a entrada de dados
    If (KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab) And txtPercentual.Text <> "" Then
        KeyAscii = 0
        sbAtribuir_Percentual_Grid
    ElseIf (KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab) And txtPercentual.Text = "" Then
        KeyAscii = 0
        txtPercentual.Visible = False
        flxProlaboristas.TextMatrix(mLinhaFocus, 5) = ""
    ElseIf KeyAscii = vbKeyEscape Then
        KeyAscii = 0
        txtPercentual.Visible = False
    End If

End Sub

Private Sub txtPercentual_LostFocus()
    txtPercentual.Text = ""
    txtPercentual.Visible = False
End Sub

Private Sub txtPercentual_Validate(Cancel As Boolean)
    Dim strMensagem As String
    If (txtPercentual.Text = "") Then
        Exit Sub
    End If
    sbAtribuir_Percentual_Grid

End Sub

Private Sub sbAtribuir_Percentual_Grid()

    flxProlaboristas.TextMatrix(mLinhaFocus, mColunaFocus) = txtPercentual.Text

    If flxProlaboristas.Col < flxProlaboristas.Cols - 1 Then
        flxProlaboristas.Col = flxProlaboristas.Col + 1
    Else
        flxProlaboristas.Col = 1
        If flxProlaboristas.Row < flxProlaboristas.Rows - 1 Then
            flxProlaboristas.Row = flxProlaboristas.Row + 1
        End If
    End If
    txtPercentual.Text = ""
    txtPercentual.Visible = False

End Sub

