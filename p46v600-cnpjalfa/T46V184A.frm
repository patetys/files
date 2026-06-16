VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmT46V184A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inclusão de Beneficiários"
   ClientHeight    =   6660
   ClientLeft      =   2085
   ClientTop       =   1800
   ClientWidth     =   8910
   ControlBox      =   0   'False
   Icon            =   "T46V184A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtClausulaAntiga 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1755
      Left            =   30
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   4200
      Visible         =   0   'False
      Width           =   8745
   End
   Begin VB.TextBox txtClausula 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1755
      Left            =   30
      MaxLength       =   32500
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   4200
      Width           =   8745
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   2175
      Left            =   0
      TabIndex        =   10
      Top             =   3840
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   3836
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Principal"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Cláusula Beneficiária"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cmbTpCobert 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "T46V184A.frx":000C
      Left            =   1710
      List            =   "T46V184A.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3480
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   3210
      MaxLength       =   60
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3510
      Width           =   870
   End
   Begin VB.ComboBox cmbTpPessoa 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "T46V184A.frx":0010
      Left            =   180
      List            =   "T46V184A.frx":0012
      Style           =   2  'Dropdown List
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3480
      Width           =   1395
   End
   Begin MSFlexGridLib.MSFlexGrid flxBenef 
      CausesValidation=   0   'False
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Para Inserir nova linha aperte Tecla INSERT. Para alterar o texto/valor dê DUPLO CLICK."
      Top             =   0
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   6165
      _Version        =   393216
      ForeColorFixed  =   8388608
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
   End
   Begin Threed.SSCommand cmdOk 
      Height          =   300
      Left            =   7440
      TabIndex        =   9
      ToolTipText     =   "Volta à proposta."
      Top             =   6240
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
      Left            =   6390
      TabIndex        =   8
      ToolTipText     =   "Novo bem."
      Top             =   6240
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
      Left            =   5340
      TabIndex        =   7
      ToolTipText     =   "Excluir o bem."
      Top             =   6240
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Excluir"
      ForeColor       =   8388608
   End
   Begin Threed.SSCommand cmdClausula 
      Height          =   300
      Left            =   240
      TabIndex        =   6
      ToolTipText     =   "Excluir o bem."
      Top             =   6240
      Visible         =   0   'False
      Width           =   2040
      _Version        =   65536
      _ExtentX        =   3598
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Cláusula Beneficiária"
      ForeColor       =   8388608
      Enabled         =   0   'False
   End
   Begin VB.Label lblPartic 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100,00"
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   8220
      TabIndex        =   5
      Top             =   3510
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Participação Total: "
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   6720
      TabIndex        =   4
      Top             =   3510
      Width           =   1485
   End
End
Attribute VB_Name = "frmT46V184A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbdTab_Ped_Dizer_Pad As ADODB.Connection
Private mrsTab_Ped_Dizer_Pad As clsYasRecordSet
Private mbdTab_Ped_Loc_Pessoas As ADODB.Connection
Private mrsTab_Ped_Loc_Pessoas As clsYasRecordSet
Private mTxtHabilitado As Boolean
Private lMsgPadrao() As String
Private lLinColFixa As Boolean
Private mobjstcPed As stcA46V702B    'Objeto da classe de estrutura stcA46V702B (TAB_PED).
Public lNum_Apol_An As Double
Public lNum_Item As Integer
Public lTipoPessoa As Integer
Public lTipoBem As Integer
Private ClausulaAntiga As Boolean

Private Sub Form_Activate()
    Dim i As Integer

    i = 1
    Do While i < flxBenef.Rows
        flxBenef.RowHeight(i) = cmbTpPessoa.Height
        i = i + 1
    Loop

End Sub

Private Sub Form_Load()
    Dim lI As Long
    Dim lSelect As String
    Dim pMensagem As String
    Dim lPartic As Currency
    Dim lCpfCnpj As String

    Call gM46V999.gpCenForm(Me)
    cmbTpCobert.Visible = False
    cmbTpPessoa.Visible = False
    Text1.Visible = False

    With cmbTpCobert
        .AddItem "Prédio e Conteúdo"
        .ItemData(.NewIndex) = 0
        .AddItem "Prédio"
        .ItemData(.NewIndex) = 1
        .AddItem "Conteúdo"
        .ItemData(.NewIndex) = 2
    End With

    With cmbTpPessoa
        .AddItem "Jurídica"
        .ItemData(.NewIndex) = 0
        .AddItem "Física"
        .ItemData(.NewIndex) = 1
        .AddItem "Instituição Financeira"
        .ItemData(.NewIndex) = 2
    End With

    With flxBenef
        .FixedRows = 1
        .FixedCols = 1
        .Rows = 2
        .Cols = 8
        .TextMatrix(0, 0) = "Nº"
        .TextMatrix(0, 1) = "Nome"
        .TextMatrix(0, 2) = "Tipo Pessoa"
        .TextMatrix(0, 3) = "CPF/CNPJ"
        .TextMatrix(0, 4) = "Tipo Cobertura"
        .TextMatrix(0, 5) = "Participação"
        .TextMatrix(0, 6) = ""
        .TextMatrix(0, 7) = ""
        .ColWidth(6) = 0
        .ColWidth(7) = 0
        .TextMatrix(1, 0) = .Rows - 1
        .TextMatrix(1, 1) = ""
        .TextMatrix(1, 2) = ""
        .TextMatrix(1, 3) = ""
        .TextMatrix(1, 4) = ""
        .TextMatrix(1, 5) = ""
        .TextMatrix(1, 6) = ""
        .TextMatrix(1, 7) = ""
        .ColWidth(0) = 350
        .ColWidth(1) = 2300
        .ColWidth(2) = 1700
        .ColWidth(3) = 1600
        .ColWidth(4) = 1500
        .ColWidth(5) = 1000

        .Row = 0: .Col = 0
        .Row = 1
    End With

    flxBenef.RowHeight(1) = cmbTpPessoa.Height
    lblPartic.Caption = FormatNumber(0, 2)

    'Procurar dizer na TAB_PED_DIZER_PAD com o código 8500
'    Set mobjstcPed = New stcA46V702B
'    Set mobjstcPed.M46V999 = gM46V999
'    If Not frmT46V012A.mfProcuraDizer(mobjstcPed, 8500, pMensagem) Then
'        End
'    Else
'        ReDim lMsgPadrao(20)
'        lMsgPadrao(0) = mobjstcPed.ColecaoTextoEspec.Item(1).Espec_lin1
'        lMsgPadrao(1) = mobjstcPed.ColecaoTextoEspec.Item(1).Espec_lin2
'        lMsgPadrao(2) = mobjstcPed.ColecaoTextoEspec.Item(1).Espec_lin3
'        lMsgPadrao(3) = mobjstcPed.ColecaoTextoEspec.Item(1).Espec_lin4
'        lMsgPadrao(4) = mobjstcPed.ColecaoTextoEspec.Item(1).Espec_lin5
'        lMsgPadrao(5) = mobjstcPed.ColecaoTextoEspec.Item(1).Espec_lin6
'        lMsgPadrao(6) = mobjstcPed.ColecaoTextoEspec.Item(1).Espec_lin7
'        lMsgPadrao(7) = mobjstcPed.ColecaoTextoEspec.Item(1).Espec_lin8
'        lMsgPadrao(8) = mobjstcPed.ColecaoTextoEspec.Item(1).Espec_lin9
'        lMsgPadrao(9) = mobjstcPed.ColecaoTextoEspec.Item(1).Espec_lin10
'        lMsgPadrao(10) = mobjstcPed.ColecaoTextoEspec.Item(1).Espec_lin11
'        lMsgPadrao(11) = mobjstcPed.ColecaoTextoEspec.Item(1).Espec_lin12
'        lMsgPadrao(12) = mobjstcPed.ColecaoTextoEspec.Item(1).Espec_lin13
'        lMsgPadrao(13) = mobjstcPed.ColecaoTextoEspec.Item(1).Espec_lin14
'        lMsgPadrao(14) = mobjstcPed.ColecaoTextoEspec.Item(1).Espec_lin15
'        lMsgPadrao(15) = mobjstcPed.ColecaoTextoEspec.Item(1).Espec_lin16
'        lMsgPadrao(16) = mobjstcPed.ColecaoTextoEspec.Item(1).Espec_lin17
'        lMsgPadrao(17) = mobjstcPed.ColecaoTextoEspec.Item(1).Espec_lin18
'        lMsgPadrao(18) = mobjstcPed.ColecaoTextoEspec.Item(1).Espec_lin19
'        lMsgPadrao(19) = mobjstcPed.ColecaoTextoEspec.Item(1).Espec_lin20
'    End If
'
'    gCarga = True    'Para não entrar no txtTextoEspec_Change.
'
'    With frmT46V155A.txtTextoEspec
'        If Len(Replace(.Text, vbCrLf, "")) = 0 Then
'            .Text = ""
'            For lI = 0 To UBound(lMsgPadrao) - 1
'                .Text = .Text & lMsgPadrao(lI) & vbCrLf
'            Next
'        End If
'    End With
'
'    gCarga = False
SAIDA:
    GoSub FechaObj
    Exit Sub
FechaObj:
    Call gM46V999.gpFecha2(mbdTab_Ped_Dizer_Pad)
    Set mrsTab_Ped_Dizer_Pad = Nothing
    Return
End Sub

Private Sub flxBenef_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Text1_LostFocus
    Call cmbTpPessoa_LostFocus
    Call cmbTpCobert_LostFocus

    cmbTpCobert.Visible = False
    cmbTpPessoa.Visible = False
    Text1.Visible = False

    Select Case flxBenef.Col
    Case 1, 3, 5
        Text1.Move flxBenef.Left + flxBenef.CellLeft, _
                   flxBenef.Top + flxBenef.CellTop, _
                   flxBenef.CellWidth, flxBenef.CellHeight
        Text1.Tag = flxBenef.Row & "," & flxBenef.Col
        Text1.Text = flxBenef.Text
        Text1.ZOrder
        Text1.Visible = True

    Case 2
        cmbTpPessoa.Move flxBenef.Left + flxBenef.CellLeft, _
                         flxBenef.Top + flxBenef.CellTop, _
                         flxBenef.CellWidth
        cmbTpPessoa.Tag = flxBenef.Row & "," & flxBenef.Col
        If Len(Trim(flxBenef.Text)) > 0 Then
            cmbTpPessoa.Text = flxBenef.Text
        End If
        cmbTpPessoa.ZOrder
        cmbTpPessoa.Visible = True

    Case 4
        cmbTpCobert.Move flxBenef.Left + flxBenef.CellLeft, _
                         flxBenef.Top + flxBenef.CellTop, _
                         flxBenef.CellWidth
        cmbTpCobert.Tag = flxBenef.Row & "," & flxBenef.Col
        If Len(Trim(flxBenef.Text)) > 0 Then
            cmbTpCobert.Text = flxBenef.Text
        End If
        cmbTpCobert.ZOrder
        cmbTpCobert.Visible = True
    End Select
End Sub

Private Sub TabStrip1_Click()

    With TabStrip1
        If (TabStrip1.SelectedItem.index = 1) Then
            txtClausula.Visible = True
            txtClausulaAntiga.Visible = False
        Else
            txtClausula.Visible = False
            txtClausulaAntiga.Visible = True
        End If
    End With

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii

    Case 8    'BackSpace
        Exit Sub

    Case 13    'Enter
        Exit Sub

    Case 22    'Ctrl+V
        If flxBenef.Col = 1 Then    'Nome
            Text1.Text = Left(Trim(Clipboard.GetText), 55)
        End If

        If flxBenef.Col = 3 Then    'CNPJ/CPF
            Dim lcopy As String

            lcopy = Replace(Replace(Replace(Clipboard.GetText, ".", Empty), "-", Empty), "/", Empty)
            If IsNumeric(lcopy) = True Then

                If UCase(flxBenef.TextMatrix(flxBenef.Row, 2)) = "FÍSICA" Then    'CPF
                    If Len(lcopy) > 11 Then
                        Beep
                    Else
                        Text1.Text = lcopy
                    End If
                Else
                    If Len(lcopy) > 14 Then
                        Beep
                    Else
                        Text1.Text = lcopy
                    End If
                End If
            End If

        Else
            If flxBenef.Col = 5 Then    'Participação

            End If
        End If

    Case 44    'Vírgula
        If flxBenef.Col = 5 Then    'Participação
            If InStr(Text1.Text, Chr(44)) > 0 Then
                KeyAscii = 0
            Else
                Text1.Text = Text1.Text & Chr(KeyAscii)
                End If
        End If

    Case 48 To 57    'Zero a nove
        If flxBenef.Col = 1 Then    'Nome
            If Len(Text1.Text) < 55 Then  ''2016-08-24: Old=60
                Text1.Text = Text1.Text & Chr(KeyAscii)
            Else
                Beep
            End If

        Else
            If flxBenef.Col = 3 Then    'CNPJ/CPF
                If UCase(flxBenef.TextMatrix(flxBenef.Row, 2)) = "FÍSICA" Then    'CPF
                    If Len(Text1.Text) < 11 Then
                        Text1.Text = Text1.Text & Chr(KeyAscii)
                    Else
                        Beep
                    End If
                Else
                    If Len(Text1.Text) < 14 Then    'CNPJ
                        Text1.Text = Text1.Text & Chr(KeyAscii)
                    Else
                        Beep
                    End If
                End If

            Else    'Participação
                If Len(Text1.Text) = 0 Then
                    Text1.Text = Chr(KeyAscii)
                Else
                    Dim lCasasAposVirgula As Integer

                    If InStr(Text1.Text, Chr(44)) = 0 Then
                        lCasasAposVirgula = 0
                    Else
                        lCasasAposVirgula = Len(Text1.Text) + 1 - InStr(Text1.Text, Chr(44))
                    End If

                    If CDbl(Text1.Text & Chr(KeyAscii)) <= 100# And lCasasAposVirgula < 3 Then
                        Text1.Text = Text1.Text & Chr(KeyAscii)
                    Else
                        Beep
                    End If
                End If

            End If    'if CNPJ/CPF

        End If    'if Nome

    Case 32, 33, 35 To 38, 40 To 47, 61, 65 To 93, 95, 97 To 123, 126
        If flxBenef.ColSel = 1 Then
            If Len(Text1.Text) < 55 Then  ''2016-08-24: Old=60
                Text1.Text = Text1.Text & Chr(KeyAscii)
            Else
                Beep
            End If
        End If

    Case Else
        Beep
    End Select

    KeyAscii = 0
    Text1.SelStart = Len(Text1.Text)
End Sub

Private Sub cmdNovo_Click()
    cmbTpCobert.Visible = False
    cmbTpPessoa.Visible = False
    Text1.Visible = False

    With flxBenef
        If .Rows < 11 Then    'Limitado a 10 beneficiarios conforme e-mail de 18/03/2014 da Alice.
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = .Rows - 1
            .TextMatrix(.Rows - 1, 1) = ""
            .TextMatrix(.Rows - 1, 2) = ""
            .TextMatrix(.Rows - 1, 3) = ""
            .TextMatrix(.Rows - 1, 4) = ""
            .TextMatrix(.Rows - 1, 5) = ""
            .TextMatrix(.Rows - 1, 6) = ""
            .TextMatrix(.Rows - 1, 7) = ""
            If .Enabled And .Visible Then
                .SetFocus
            End If
            .Row = .Rows - 1
            .RowSel = .Rows - 1
            .ColSel = 1
            .Col = 1
            .RowHeight(.Row) = cmbTpPessoa.Height
        End If
    End With
End Sub

Private Sub cmdExcluir_Click()
    Dim lI As Integer
    Dim lRowSel As Integer

    cmbTpCobert.Visible = False
    cmbTpPessoa.Visible = False
    Text1.Visible = False

    If flxBenef.Rows > 2 Then
        lRowSel = flxBenef.RowSel
        If flxBenef.RowSel <> 0 Then
            flxBenef.RemoveItem flxBenef.RowSel

            For lI = 1 To flxBenef.Rows - 1
                flxBenef.TextMatrix(lI, 0) = lI
            Next
        End If
    Else
        flxBenef.Row = 1
        flxBenef.TextMatrix(1, 1) = Empty
        flxBenef.TextMatrix(1, 2) = Empty
        flxBenef.TextMatrix(1, 3) = Empty
        flxBenef.TextMatrix(1, 4) = Empty
        flxBenef.TextMatrix(1, 5) = Empty
    End If

    Call mpRecalcPartic
End Sub

Private Sub cmdOk_Click()
    Dim lI As Integer
    Dim lLin As Integer
    Dim lcol As Integer
    Dim lTemBenef As Boolean

    '    Const lWildCard = "(*)"
    '
    '    gCarga = True    'Para não entrar no txtTextoEspec_Change.
    '
    '    With frmT46V155A.txtTextoEspec
    '        .Text = Empty
    '
    '        If Len(Join(lMsgPadrao, Empty)) > 0 Then    'Verifica se o array tem conteúdo.
    '            For lI = 0 To UBound(lMsgPadrao) - 1
    '                If Trim(lMsgPadrao(lI)) <> lWildCard Then
    '                    .Text = .Text & lMsgPadrao(lI) & vbCrLf
    '                Else
    '                    For lLin = 1 To flxBenef.Rows - 1
    '                        For lcol = 1 To flxBenef.Cols - 3
    '                            .Text = .Text & flxBenef.TextMatrix(lLin, lcol)
    '
    '                            If lcol < flxBenef.Cols - 3 Then
    '                                .Text = .Text & ", "
    '                            Else
    '                                .Text = .Text & vbCrLf
    '                            End If
    '                        Next
    '                    Next
    '                End If
    '            Next
    '        End If
    '    End With
    '
    '    frmT46V155A.txtTextoEspec.Enabled = mTxtHabilitado    'Recupera situação anterior.

    With frmT46V184A
        lTemBenef = False

        For lLin = 1 To .flxBenef.Rows - 1
            For lcol = 1 To .flxBenef.Cols - 1
                If Trim(.flxBenef.TextMatrix(lLin, lcol)) <> Empty Then
                    lTemBenef = True
                    Exit For
                End If
            Next
        Next
    End With

    If lTemBenef = True Then
        frmT46V101A.optBeneficiario(0).Value = True
        frmT46V101A.optBeneficiario(1).Value = False
    Else
        frmT46V101A.optBeneficiario(0).Value = False
        frmT46V101A.optBeneficiario(1).Value = True
    End If

    If gM46V999.gfDisponivel("frmT46V184A") Then
        Me.Hide
    End If

End Sub

Private Sub cmdClausula_Click()
    Me.Hide
    mTxtHabilitado = frmT46V155A.txtTextoEspec.Enabled    'Guarda situação atual.
    frmT46V155A.txtTextoEspec.Enabled = False    'Não permite digitar.
    frmT46V155A.Show vbModal, Me
End Sub

Private Sub mpRecalcPartic()
    Dim lI As Integer
    Dim lPartic As Double

    lPartic = 0#

    For lI = 1 To flxBenef.Rows - 1
        If IsNumeric(flxBenef.TextMatrix(lI, 5)) Then
            lPartic = lPartic + gM46V999.CDblx(flxBenef.TextMatrix(lI, 5))
        End If
    Next

    lblPartic.Caption = FormatNumber(lPartic, 2)
End Sub

Private Sub Text1_LostFocus()
    Dim lcol As Integer
    Dim lrow As Integer

    If Len(Text1.Tag) Then
        lrow = Left(Text1.Tag, InStr(Text1.Tag, ","))
        lcol = Right(Text1.Tag, InStr(Text1.Tag, ",") - 1)

        If lrow > flxBenef.Rows - 1 Then Exit Sub

        Select Case lcol
        Case 1
            
            flxBenef.TextMatrix(lrow, lcol) = Text1.Text ''(Text1.Text + space(60))

        Case 3
            Dim lCpfCnpj As String
            lCpfCnpj = Text1.Text
            lCpfCnpj = Replace(Replace(Replace(lCpfCnpj, ".", Empty), "-", Empty), "/", Empty)
            If UCase(flxBenef.TextMatrix(lrow, 2)) = "FÍSICA" Then
                Call gM46V999.gpForCPF(1, lCpfCnpj, lCpfCnpj)
            Else
                Call gM46V999.gpForCNPJ(1, lCpfCnpj, lCpfCnpj)
            End If

            flxBenef.TextMatrix(lrow, lcol) = lCpfCnpj
        Case 5
            If IsNumeric(Text1.Text) = True Then
                flxBenef.TextMatrix(lrow, lcol) = FormatNumber(Text1.Text, 2)
            Else
                flxBenef.TextMatrix(lrow, lcol) = FormatNumber("0", 2)
            End If
            Call mpRecalcPartic
        End Select

    End If
    Text1.Visible = False
End Sub

Private Sub cmbTpPessoa_LostFocus()
    If cmbTpPessoa.Visible = True And Len(cmbTpPessoa.Tag) > 0 Then
        If flxBenef.TextMatrix(Left(cmbTpPessoa.Tag, InStr(cmbTpPessoa.Tag, ",")), Right(cmbTpPessoa.Tag, InStr(cmbTpPessoa.Tag, ",") - 1)) <> cmbTpPessoa.Text Then
            flxBenef.TextMatrix(Left(cmbTpPessoa.Tag, InStr(cmbTpPessoa.Tag, ",")), 3) = Empty
        End If
        flxBenef.TextMatrix(Left(cmbTpPessoa.Tag, InStr(cmbTpPessoa.Tag, ",")), Right(cmbTpPessoa.Tag, InStr(cmbTpPessoa.Tag, ",") - 1)) = cmbTpPessoa.Text
    End If
End Sub

Private Sub cmbTpCobert_LostFocus()
    If cmbTpCobert.Visible = True And Len(cmbTpCobert.Tag) Then
        flxBenef.TextMatrix(Left(cmbTpCobert.Tag, InStr(cmbTpCobert.Tag, ",")), Right(cmbTpCobert.Tag, InStr(cmbTpCobert.Tag, ",") - 1)) = cmbTpCobert.Text
    End If
End Sub

Public Sub CarregarBeneficiariosAntigos()
    Dim lI As Long
    Dim pMensagem As String
    Dim lSaida() As Variant         'Resultado da busca de Beneficiarios.
    Dim lQtdeRegistros As Double    'Quantidade de beneficiarios.


    If (lNum_Apol_An <> 0 And lNum_Item <> 0) Then
        If gM46V999.gfObtBuscaBeneficiariosAntigos(lNum_Apol_An, lNum_Item, lSaida, pMensagem) = False Then
            Exit Sub
        End If

        lQtdeRegistros = UBound(lSaida)
        For lI = 1 To lQtdeRegistros
            txtClausulaAntiga.Text = lSaida(lI, 1)
        Next lI

    End If
End Sub

Public Sub CarregarNovaClausulaBeneficiaria()
    Dim lI As Long
    Dim pMensagem As String
    Dim lSaida() As Variant         'Resultado da busca de Beneficiarios.
    Dim lQtdeRegistros As Double    'Quantidade de beneficiarios.

    If gM46V999.gfObtBuscaNovaClausulaBeneficiaria(lTipoPessoa, lTipoBem, lSaida, pMensagem) = False Then
        Call gM46V999.gpGraLog(1, pMensagem)
        Exit Sub
    End If

    lQtdeRegistros = UBound(lSaida)
    txtClausula.Text = Empty
    For lI = 1 To lQtdeRegistros
        txtClausula.Text = txtClausula.Text & lSaida(lI, 1)
    Next lI

End Sub
