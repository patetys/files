VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmT46V168A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LMI"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9090
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   9090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid flxLMI 
      Height          =   3375
      Left            =   60
      TabIndex        =   0
      ToolTipText     =   "Para Inserir nova linha aperte Tecla INSERT"
      Top             =   30
      Width           =   9030
      _ExtentX        =   15928
      _ExtentY        =   5953
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin Threed.SSCommand cmdNovo 
      Height          =   300
      Left            =   210
      TabIndex        =   1
      ToolTipText     =   "Novo bem."
      Top             =   3855
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
      Left            =   1440
      TabIndex        =   2
      ToolTipText     =   "Excluir o bem."
      Top             =   3855
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Excluir"
      ForeColor       =   8388608
   End
   Begin Threed.SSCommand cmdOk 
      Height          =   300
      Left            =   3795
      TabIndex        =   3
      ToolTipText     =   "Volta à proposta."
      Top             =   3885
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Ok"
      ForeColor       =   8388608
   End
End
Attribute VB_Name = "frmT46V168A"
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

Private Sub flxLMI_LostFocus()
    With flxLMI
        If .ColSel = 1 Then
            Call flxLMI_LeaveCell
        End If
    End With
End Sub

Private Sub Form_Load()

    Call gM46V999.gpCenForm(Me)
    With flxLMI
        .FixedRows = 1
        .FixedCols = 1
        .Rows = 2
        .Cols = 7
        .TextMatrix(0, 0) = "Nº"
        .TextMatrix(0, 1) = "CEP"
        .TextMatrix(0, 2) = "LOGRADOURO"
        .TextMatrix(0, 3) = "BAIRRO"
        .TextMatrix(0, 4) = "CIDADE"
        .TextMatrix(0, 5) = "UF"
        .TextMatrix(0, 6) = "CEP INVISÍVEL"
        .ColWidth(0) = 360
        .ColWidth(1) = 900
        .ColWidth(2) = 3500
        .ColWidth(3) = 1600
        .ColWidth(4) = 4000
        .ColWidth(5) = 360
        .ColWidth(6) = 0
        .TextMatrix(1, 0) = "001"
        .TextMatrix(1, 1) = ""
        .TextMatrix(1, 2) = ""
        .TextMatrix(1, 3) = ""
        .TextMatrix(1, 4) = ""
        .TextMatrix(1, 5) = ""
        .TextMatrix(1, 6) = ""
        .Row = 1
        .ColAlignment(2) = flexAlignLeftTop
    End With

    mLinhaFocus = 1

End Sub
Private Sub Form_Activate()
    frmT46V100A.panForm.Caption = Mid$(Me.Name, 4)
Finalizar:
    '    If gM46V999.gDesabilitarForms Then
    '        If Me.Tag <> "Desabilitado" Then
    '            DesabilitaControles Me
    '        End If
    '    Else
    '        If Me.Tag = "Desabilitado" Then
    '            HabilitaControles Me
    '        End If
    '    End If

End Sub
Private Sub flxLMI_KeyDown(KeyCode As Integer, Shift As Integer)
    With flxLMI
        If .RowSel > 0 And .ColSel > 1 Then
            Select Case KeyCode
            Case 46  'delete
                .Row = .RowSel
                .Col = .ColSel
                .Refresh
                .TextMatrix(.RowSel, .ColSel) = ""
            Case 45
                Call cmdNovo_Click
            End Select
        End If
    End With
    Call gpHabSalvar
End Sub
Private Sub flxLMI_Click()
    With flxLMI
        mColunaFocus = .ColSel
        mLinhaFocus = .RowSel
    End With

End Sub
Private Sub flxLMI_KeyPress(KeyAscii As Integer)
    Dim lTamanho As Long
    With flxLMI
        If .RowSel > 0 And .ColSel > 0 Then
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
            Case 32, 33, 35 To 38, 40 To 47, 48 To 57, 61, 65 To 93, 95, 97 To 123, 126
                If .ColSel <> 0 Then
                    If .ColSel = 5 Then    'UF
                        lTamanho = 2
                    ElseIf .ColSel = 4 Then    'Cidade
                        lTamanho = 40
                    ElseIf .ColSel = 3 Then    'Bairro
                        lTamanho = 10
                    ElseIf .ColSel = 2 Then    'Logradouro
                        lTamanho = 30
                    ElseIf .ColSel = 1 Then    'Cep
                        lTamanho = 9
                    End If
                    If Len(.TextMatrix(.RowSel, .ColSel)) < lTamanho Then
                        .TextMatrix(.RowSel, .ColSel) = .TextMatrix(.RowSel, .ColSel) & Chr(KeyAscii)
                    Else
                        Beep
                    End If
                End If
            Case 22    'deve o colar.
                If .ColSel = 5 Then    'UF
                    lTamanho = 2
                ElseIf .ColSel = 4 Then    'Cidade
                    lTamanho = 40
                ElseIf .ColSel = 3 Then    'Bairro
                    lTamanho = 10
                ElseIf .ColSel = 2 Then    'Logradouro
                    lTamanho = 30
                ElseIf .ColSel = 1 Then    'Cep
                    lTamanho = 9
                End If
                .TextMatrix(.RowSel, .ColSel) = .TextMatrix(.RowSel, .ColSel) & Clipboard.GetText
                If Len(.TextMatrix(.RowSel, .ColSel)) > lTamanho Then
                    MsgBox "FGE0040 - O tamanho do campo colado é maior do que o permitido. (" & lTamanho & ")"
                    .TextMatrix(.RowSel, .ColSel) = ""
                End If
            End Select
            Call gpHabSalvar
        End If
    End With

End Sub
Private Sub flxLMI_LeaveCell()
    Dim lCEP As String    'CEP.
    Dim lMensagem As String    'Mensagem.
    Dim lPos As Long


    With flxLMI
        If .ColSel = 1 Then
            'Se CEP não informado, encerra processo.
            If Not gM46V999.gfPreenchido(.TextMatrix(.RowSel, 1)) Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            '6. Verifica se o CEP está preenchido corretamente.
            .TextMatrix(.RowSel, 1) = Trim$(.TextMatrix(.RowSel, 1))
            Select Case Len(.TextMatrix(.RowSel, 1))
            Case Is < 8
                lPos = InStr(1, .TextMatrix(.RowSel, 1), "-")
                If lPos > 6 Then
                    .TextMatrix(.RowSel, 1) = .TextMatrix(.RowSel, 1)
                ElseIf lPos < 6 And lPos <> 0 Then
                    .TextMatrix(.RowSel, 1) = Format(Left(.TextMatrix(.RowSel, 1), lPos - 1), "00000") & "-" & Format(Left(Mid(.TextMatrix(.RowSel, 1), lPos + 1), 3), "000")
                Else
                    .TextMatrix(.RowSel, 1) = Format(Format(.TextMatrix(.RowSel, 1), "00000000"), "00000-000")
                End If
            Case 8
                lPos = InStr(1, .TextMatrix(.RowSel, 1), "-")
                If lPos > 6 Then
                    .TextMatrix(.RowSel, 1) = .TextMatrix(.RowSel, 1)
                ElseIf lPos <= 6 And lPos <> 0 Then
                    .TextMatrix(.RowSel, 1) = Format(Left(.TextMatrix(.RowSel, 1), lPos - 1), "00000") & Left(Mid(.TextMatrix(.RowSel, 1), lPos + 1) & "000", 3)
                End If

                If Not IsNumeric(.TextMatrix(.RowSel, 1)) Then
                    MsgBox "FGE0039 - CEP inválido (não numérico).", vbExclamation
                    .TextMatrix(.RowSel, 1) = ""
                    If .Visible = True And .Enabled = True Then
                        .Row = .RowSel
                        .Col = 1
                    End If
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                .TextMatrix(.RowSel, 1) = Mid$(.TextMatrix(.RowSel, 1), 1, 5) & "-" & Mid$(.TextMatrix(.RowSel, 1), 6, 3)
                lCEP = .TextMatrix(.RowSel, 1)
            Case 9
                If Not IsNumeric(Left$(.TextMatrix(.RowSel, 1), 5)) Then
                    MsgBox "FGE0039 - CEP inválido " & _
                           "(fora do formato nnnnn-nnn - primeira parte do CEP não numérica).", vbExclamation
                    .TextMatrix(.RowSel, 1) = ""
                    If .Visible = True And .Enabled = True Then
                        .Row = .RowSel
                        .Col = 1
                    End If
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                If Not IsNumeric(Right$(.TextMatrix(.RowSel, 1), 3)) Then
                    MsgBox "FGE0039 - CEP inválido " & _
                           "(fora do formato nnnnn-nnn - segunda parte do CEP não numérica).", vbExclamation
                    .TextMatrix(.RowSel, 1) = ""
                    If .Visible = True And .Enabled = True Then
                        .Row = .RowSel
                        .Col = 1
                    End If
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                If Mid$(.TextMatrix(.RowSel, 1), 6, 1) <> "-" Then
                    MsgBox "FGE0039 - CEP inválido " & _
                           "(fora do formato nnnnn-nnn - sem o traço).", vbExclamation
                    .TextMatrix(.RowSel, 1) = ""
                    If .Visible = True And .Enabled = True Then
                        .Row = .RowSel
                        .Col = 1
                    End If
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            End Select
            If Right$(.TextMatrix(.RowSel, 1), 3) = "999" Then
                MsgBox "FGE0039 - " & _
                       "CEP inválido (final inválido - igual a 999).", vbExclamation
                .TextMatrix(.RowSel, 1) = ""
                If .Visible = True And .Enabled = True Then
                    .Row = .RowSel
                    .Col = 1
                End If
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            If Not IsNumeric(Replace(.TextMatrix(.RowSel, 1), "-", "")) Then
                .TextMatrix(.RowSel, 1) = "0"
            End If

            'Preenche o endereço
            If .TextMatrix(.RowSel, 6) <> .TextMatrix(.RowSel, 1) Or _
               (.TextMatrix(.RowSel, 6) = .TextMatrix(.RowSel, 1) And (.TextMatrix(.RowSel, 2) = "" Or .TextMatrix(.RowSel, 3) = "" Or .TextMatrix(.RowSel, 4) = "" Or .TextMatrix(.RowSel, 5) = "")) Then
                Call mpPreencheEndereco
            End If
        End If
    End With
    Screen.MousePointer = vbDefault

End Sub
Private Sub mpPreencheEndereco()
    Dim lTipLogradouro As String
    Dim lNomLogradouro As String
    Dim lBairro As String
    Dim lCidade As String
    Dim lUF As String
    Dim lMsg As String

    With flxLMI
        If gfBuscaPorCEP(.TextMatrix(.RowSel, 1), lTipLogradouro, lNomLogradouro, lBairro, lCidade, lUF, lMsg) = False Then
            Call gM46V999.gpGraLog(1, lMsg, vbInformation)
        Else
            If gM46V999.gfPreenchido(lTipLogradouro) = True Then
                .TextMatrix(.RowSel, 2) = Trim(lTipLogradouro) & " " & Trim(lNomLogradouro) & ", "
                .TextMatrix(.RowSel, 3) = Left(Trim(lBairro), 10)
            End If
            .TextMatrix(.RowSel, 4) = Trim(lCidade)
            .TextMatrix(.RowSel, 5) = lUF
            .TextMatrix(.RowSel, 6) = .TextMatrix(.RowSel, 1)
        End If
    End With
End Sub
Public Function gfBuscaPorCEP(ByVal pCEP As String, _
                              ByRef pTipLogradouro As String, _
                              ByRef pNomLogradouro As String, _
                              ByRef pBairro As String, _
                              ByRef pCidade As String, _
                              ByRef pUF As String, _
                              ByRef pMensagem As String) As Boolean

'Dim lCN_BaseCEP         As ADODB.Connection     'Conexão com a Base de CEP
    Dim lbdP0043000 As ADODB.Connection    'Conexão com a Base de CEP
    Dim lRS_BaseCEP As clsYasRecordSet
    Dim lSelect As String
    Dim lMensagem As String     'Mensagem.


    Err = 0
    lMensagem = ""
    On Error GoTo Func_Erro

    Screen.MousePointer = vbHourglass

    If Not gM46V999.gfPreenchido(pCEP) Then
        gfBuscaPorCEP = True
        GoTo Func_Saida
    End If

    If Len(Trim(pCEP)) = 8 Then
        If IsNumeric(pCEP) Then
            pCEP = Left$(pCEP, 5) & "-" & Right$(pCEP, 3)
        End If
    End If

    '2. Consiste CEP.
    If Not IsNumeric(Mid$(pCEP, 1, 5)) Or Not IsNumeric(Mid$(pCEP, 7, 3)) Then
        lMensagem = "FGE2079 - CEP inválido (endereço do segurado - " & pCEP & _
                  " - não numérico - fora do formato nnnnn-nnn)."
        GoTo Func_Erro
    End If

    If Mid$(pCEP, 6, 1) <> "-" Then
        lMensagem = "FGE2079 - CEP inválido (endereço do segurado - " & pCEP & _
                  " - não tem traço na posição 6 - fora do formato nnnnn-nnn)."
        GoTo Func_Erro
    End If

    '3. Obtém CEP.

    If gM46V999.gfAbrBasDados("P0043000", gM46V999.gTipP0043000, gM46V999.gEndP0043000, lbdP0043000, lMensagem) = False Then
        GoTo Func_Erro
    End If

    'b) Obtém registro.
    lSelect = "SELECT * FROM P0043000 WHERE NUM_CEP = '" & Left$(pCEP, 5) & Right$(pCEP, 3) & "'"
    If gM46V999.gfObtRegistro(lbdP0043000, lSelect, lRS_BaseCEP, lMensagem) = False Then
        GoTo Func_Erro
    End If

    If lRS_BaseCEP.EOF = True Then
        lMensagem = "FNL0048 - CEP não localizado (" & pCEP & ")."
        GoTo Func_Erro
    End If

    'c) Retorna o endereço encontrado
    If gM46V999.gfPreenchido(lRS_BaseCEP!TIP_LOGR) Then
        pTipLogradouro = lRS_BaseCEP!TIP_LOGR
    Else
        pTipLogradouro = ""
    End If
    If gM46V999.gfPreenchido(lRS_BaseCEP!Nom_Logr) Then
        pNomLogradouro = lRS_BaseCEP!Nom_Logr
    Else
        pNomLogradouro = ""
    End If
    If gM46V999.gfPreenchido(lRS_BaseCEP!Nom_Bairro) Then
        pBairro = lRS_BaseCEP!Nom_Bairro
    Else
        pBairro = ""
    End If
    If gM46V999.gfPreenchido(lRS_BaseCEP!Nom_Cid) Then
        pCidade = lRS_BaseCEP!Nom_Cid
    Else
        pCidade = ""
    End If
    If gM46V999.gfPreenchido(lRS_BaseCEP!Sig_UF) Then
        pUF = lRS_BaseCEP!Sig_UF
    Else
        pUF = ""
    End If

    gfBuscaPorCEP = True
Func_Saida:
    'd) Fecha tabela e arquivo.
    Call gM46V999.gpFecha2(lbdP0043000)

    Err = 0
    On Error GoTo 0

    Screen.MousePointer = vbDefault
    Exit Function

Func_Erro:
    gfBuscaPorCEP = False
    If gM46V999.gfPreenchido(lMensagem) = True Then
        pMensagem = lMensagem
    Else
        pMensagem = "gfBuscaPorCEP - Erro " & Err & "  " & Error
    End If
    GoTo Func_Saida

End Function


Private Sub cmdExcluir_Click()
    Dim lRowSel As Long    'Linha selecionada.
    Dim lI As Long              'Utilizada no for...loop.

    If Not flxLMI.Enabled Then Exit Sub

    If flxLMI.Rows > 2 Then
        lRowSel = flxLMI.RowSel
        If flxLMI.RowSel <> 0 Then
            flxLMI.RemoveItem flxLMI.RowSel
        End If
        For lI = lRowSel To flxLMI.Rows - 1
            flxLMI.TextMatrix(lI, 0) = Format(lI, "000")
        Next lI
        mLinhaFocus = mLinhaFocus - 1
    Else
        lI = 1
        flxLMI.Row = lI
        flxLMI.TextMatrix(lI, 0) = Format(lI, "000")
        flxLMI.TextMatrix(lI, 1) = ""
        flxLMI.TextMatrix(lI, 2) = ""
        flxLMI.TextMatrix(lI, 3) = ""
        flxLMI.TextMatrix(lI, 4) = ""
        flxLMI.TextMatrix(lI, 5) = ""
        flxLMI.TextMatrix(lI, 6) = ""
    End If
    Call gpHabSalvar
End Sub

Private Sub cmdNovo_Click()
    If Not flxLMI.Enabled Then Exit Sub
    With flxLMI
        If .Rows < 999 Then
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
            flxLMI_Click
        End If
    End With

End Sub

Private Sub cmdOK_Click()
    Dim lI As Long    'Utilizada no for...loop.
'Verificar as informações digitadas.
    For lI = 1 To flxLMI.Rows - 1
        If lI >= flxLMI.Rows Then Exit For
        If flxLMI.TextMatrix(lI, 1) = "" And flxLMI.TextMatrix(lI, 2) = "" And flxLMI.TextMatrix(lI, 3) = "" And flxLMI.TextMatrix(lI, 4) = "" Then
            If flxLMI.Rows = 2 Then
                gCarga = True
                frmT46V143B.CkEndLMI.Value = False
                gCarga = False
                Exit For
            Else
                flxLMI.RemoveItem lI
                lI = lI - 1
            End If
        Else
            If flxLMI.TextMatrix(lI, 1) <> "" Or flxLMI.TextMatrix(lI, 2) <> "" Or flxLMI.TextMatrix(lI, 3) <> "" Or flxLMI.TextMatrix(lI, 4) <> "" Then
                If flxLMI.TextMatrix(lI, 1) = "" Then
                    MsgBox "FNP0001 - Favor digitar CEP.", vbExclamation
                    flxLMI.TextMatrix(lI, 1) = ""
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                If flxLMI.TextMatrix(lI, 2) = "" Then
                    MsgBox "FNP0001 - Favor digitar LOGRADOURO.", vbExclamation
                    flxLMI.TextMatrix(lI, 2) = ""
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                If flxLMI.TextMatrix(lI, 3) = "" Then
                    MsgBox "FNP0001 - Favor digitar BAIRRO.", vbExclamation
                    flxLMI.TextMatrix(lI, 3) = ""
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                If flxLMI.TextMatrix(lI, 4) = "" Then
                    MsgBox "FNP0002 - Favor digitar CIDADE.", vbExclamation
                    flxLMI.TextMatrix(lI, 4) = ""
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                If flxLMI.TextMatrix(lI, 5) = "" Then
                    MsgBox "FNP0003 - Favor digitar UF.", vbExclamation
                    flxLMI.TextMatrix(lI, 5) = ""
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            End If
        End If
    Next lI
    Me.Hide

End Sub

