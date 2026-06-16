VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT46V011A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Endosso"
   ClientHeight    =   4350
   ClientLeft      =   2730
   ClientTop       =   1425
   ClientWidth     =   6135
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "T46V011A.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel panEndosso 
      Height          =   3780
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5925
      _Version        =   65536
      _ExtentX        =   10451
      _ExtentY        =   6667
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      Begin VB.TextBox txtDizer 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         MaxLength       =   148
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   2400
         Width           =   5655
      End
      Begin VB.TextBox txtItem 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   4680
         MaxLength       =   4
         TabIndex        =   5
         ToolTipText     =   "Número da cotação."
         Top             =   3240
         Width           =   600
      End
      Begin VB.OptionButton optItem 
         Caption         =   "Apenas item "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   1
         Left            =   3120
         TabIndex        =   4
         Top             =   3360
         Value           =   -1  'True
         Width           =   1500
      End
      Begin VB.OptionButton optItem 
         Caption         =   "Todos os itens"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   3120
         TabIndex        =   3
         Top             =   3120
         Width           =   1500
      End
      Begin VB.TextBox txtApolice 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   2
         ToolTipText     =   "Número da apólice anterior."
         Top             =   3240
         Width           =   1100
      End
      Begin VB.Frame fmTipoEndosso 
         Caption         =   "Tipo de endosso:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   5655
         Begin VB.ComboBox cboDizer 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   240
            Width           =   5415
         End
      End
      Begin Threed.SSCommand cmdPesquisa 
         Height          =   285
         Left            =   2280
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Busca apólice."
         Top             =   3240
         Width           =   285
         _Version        =   65536
         _ExtentX        =   494
         _ExtentY        =   503
         _StockProps     =   78
         ForeColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "T46V011A.frx":000C
      End
      Begin VB.Frame fraDizer 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Visible         =   0   'False
         Width           =   5655
         Begin VB.ListBox lstDizer 
            Columns         =   3
            Height          =   1425
            Left            =   0
            TabIndex        =   13
            Top             =   120
            Width           =   5655
         End
      End
      Begin VB.Label lblX 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Apólice: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   225
         TabIndex        =   6
         Top             =   3240
         Width           =   795
      End
   End
   Begin Threed.SSCommand cmdOk 
      Height          =   300
      Left            =   120
      TabIndex        =   7
      Top             =   3960
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1764
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Ok"
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
   End
   Begin Threed.SSCommand cmdCancelar 
      Cancel          =   -1  'True
      Height          =   300
      Left            =   1320
      TabIndex        =   8
      Top             =   3960
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1764
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Cancelar"
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
   End
End
Attribute VB_Name = "frmT46V011A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mOpcaoEndosso As Integer    'Opção do endosso.
Dim mNossoNumero As String    'Nosso Numero.
Dim mArrayDizerOutros() As String
Dim mCod_Ramo As Integer
Dim mCod_Modalidade As Integer

Private Function mfGraEndosso()
'Função: processa opção de endosso de seguro Yasuda.

    Dim lMensagem As String    'Mensagem.

    mfGraEndosso = False
    On Error GoTo ControleErro

    '1. Consiste número de apólice anterior.
    If Not gM46V999.gfPreenchido(txtApolice.Text) Then
        Call gM46V999.gpGraLog(1, "FGE0069 - Número da apólice inválido (Endosso Yasuda - não preenchido).")
        If txtApolice.Visible = True And txtApolice.Enabled = True Then
            txtApolice.SetFocus
        End If
        Exit Function
    End If

    If Not IsNumeric(txtApolice.Text) Then
        Call gM46V999.gpGraLog(1, "FGE0069 - Número da apólice inválido (Endosso Yasuda - não numérico).")
        If txtApolice.Visible = True And txtApolice.Enabled = True Then
            txtApolice.SetFocus
        End If
        Exit Function
    End If

    If Len(Trim(txtApolice.Text)) < 10 Then
        txtApolice.Text = Trim$(txtApolice.Text)
        txtApolice.Text = String(10 - Len(txtApolice.Text), "0") & txtApolice.Text
    End If

    '2. Consiste número do item da apólice.
    If optItem(1) = True Then   'Apenas um item.
        If Not gM46V999.gfPreenchido(txtItem.Text) Then
            Call gM46V999.gpGraLog(1, "FGE0070 - Número do item da apólice inválido (Endosso Yasuda - não " & _
                                      "preenchido).")
            If txtItem.Visible = True And txtItem.Enabled = True Then
                txtItem.SetFocus
            End If
            Exit Function
        End If
        If Not IsNumeric(txtItem.Text) Then
            Call gM46V999.gpGraLog(1, "FGE0070 - Número do item da apólice inválido (Endosso Yasuda - não " & _
                                      "numérico).")
            If txtItem.Visible = True And txtItem.Enabled = True Then
                txtItem.SetFocus
            End If
            Exit Function
        End If
        If Len(Trim(txtItem.Text)) < 4 Then
            txtItem.Text = Trim$(txtItem.Text)
            txtItem.Text = String(4 - Len(txtItem.Text), "0") & txtItem.Text
        End If
    End If

    '3. Grava registros de cotação/proposta para Endosso Yasuda.
    'If gM46V001.gfGraRenYasuda(txtApolice.Text, txtItem.Text, mNossoNumero, _
     '        IIf(optSinistroYas(0).Value = True, e_SimNao_Sim, e_SimNao_Nao), _
     '        lMensagem) = False Then

    If gM46V001.gfGraEndossoYasuda(txtApolice.Text, txtItem.Text, mNossoNumero, _
                                   e_SimNao_Nao, cboDizer.ItemData(cboDizer.ListIndex), lMensagem, True, Nothing, "") = False Then
        If gM46V999.gfPreenchido(lMensagem) Then
            Call gM46V999.gpGraLog(1, lMensagem)
        End If
        If txtApolice.Visible = True And txtApolice.Enabled = True Then
            txtApolice.SetFocus
        End If
        Exit Function
    End If
    mfGraEndosso = True
    Exit Function
ControleErro:
    Call gM46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gM46V999.gpGraLog(0, "mfGraEndosso ERRO - " & Err & " " & Err.Description)
End Function



Private Sub cboDizer_Click()

    Call mpAjustaForm

    Select Case cboDizer.ItemData(cboDizer.ListIndex)
    Case 1  'Inclusão
        optItem(0).Value = True
        optItem(0).Enabled = False
        optItem(1).Enabled = False
        txtItem.Enabled = False
        mOpcaoEndosso = 302
        txtDizer.Text = ""
    Case 4  'Exclusão
        optItem(0).Enabled = False
        optItem(1).Enabled = True
        optItem(1).Value = True
        txtItem.Enabled = True
        mOpcaoEndosso = 303
        txtDizer.Text = ""
    Case 10    'Cancelamento
        optItem(0).Value = True
        optItem(0).Enabled = True
        optItem(1).Enabled = False
        txtItem.Enabled = False
        mOpcaoEndosso = 301
        txtDizer.Text = ""
        Call MsgBox("AGE0099- Válido somente para apólices quitadas.", vbInformation, Me.Caption)
        '        Case 12 'Alteração
        '            Me.Height = 4725
        '            cmdCancelar.Top = 3960
        '            cmdOk.Top = 3960
        '            panEndosso.Height = 3780
        '            txtApolice.Top = 3240
        '            cmdPesquisa.Top = 3240
        '            lblX(0).Top = 3240
        '            optItem(0).Top = 3120
        '            optItem(1).Top = 3360
        '            txtItem.Top = 3240
        '            txtDizer.Visible = True
        '            If txtDizer.Visible = True And txtDizer.Enabled = True Then
        '                txtDizer.SetFocus
        '            End If
        '            optItem(1).Value = True
        '            optItem(0).Enabled = False
        '            optItem(1).Enabled = True
        '            txtItem.Enabled = True
        '            mOpcaoEndosso = 300
        '            fraDizer.Visible = True
        '            Call mpCarregarDizer12
        '     '       If gM46V999.PerfilAcesso = e_Perfil_DIG_IST Then
        '     '           lstDizer.ListIndex = 1
        '     '       End If
    Case Else   'Alteração
        optItem(1).Value = True
        optItem(0).Enabled = False
        optItem(1).Enabled = True
        txtItem.Enabled = True
        mOpcaoEndosso = 300
        txtDizer.Text = ""
    End Select
End Sub

Private Sub cmdPesquisa_Click()
    Dim lNum_Apolice As String    'Número da apólice.
    Dim lNum_Item As String    'Número do item.
    Dim lCod_Ramo As Integer    'Código do Ramo
    Dim lCod_Modalidade As Integer   'Codigo do Produto / Modalidade

    Call frmT46V003A.mpInstanciarForm(lNum_Apolice, lNum_Item, lCod_Ramo, lCod_Modalidade)
    txtApolice.Text = lNum_Apolice
    txtItem.Text = lNum_Item
    mCod_Ramo = lCod_Ramo
    mCod_Modalidade = lCod_Modalidade
    If cmdOk.Visible = True And cmdOk.Enabled = True Then
        cmdOk.SetFocus
    End If

End Sub

Private Sub Form_Load()
'Dim lobjstc             As stcA46V101B          'Objeto da classe de estrutura stcA46V101B.
    Dim lbdTab_Dizer As ADODB.Connection    'Arquivo: P00Tarifa.
    Dim lMensagem As String     'Mensagem.
    Dim lrsTab_Dizer As clsYasRecordSet    'Registro : Tab_Dizer.
    Dim lLinha As Long    'Linhas.
    Dim lSelect As String

    Call gM46V999.gpCenForm(Me)  'Centraliza form.

    '1. Posiciona controles do help.

    'Inicio como Alteração
    txtItem.Enabled = True
    mOpcaoEndosso = 300
    '    gM46V999.gFormCancelado = False
    txtApolice.Text = ""
    txtItem.Text = ""

    If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTab_Dizer, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If

    lSelect = " Select * from tab_dizer "

    If gM46V999.gfObtRegistro(lbdTab_Dizer, lSelect, lrsTab_Dizer, lMensagem) = False Then
        Call gM46V999.gpFecha2(lbdTab_Dizer)
        Call gM46V999.gpGraLog(2, lMensagem)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If


    cboDizer.Clear
    Do While Not lrsTab_Dizer.EOF
        cboDizer.AddItem Format(lrsTab_Dizer("Cod_Dizer"), "00") & " - " & lrsTab_Dizer("Dsc_Dizer")
        cboDizer.ItemData(cboDizer.NewIndex) = lrsTab_Dizer("Cod_Dizer")
        lrsTab_Dizer.MoveNext
    Loop

    cboDizer.ListIndex = -1

    Call mpAjustaForm

    '    Select Case gM46V999.gTipoAcesso
    '        Case DIGITACAO
    '            'Caso seja digitação, não aparecer o botão de pesquisa
    '            cmdPesquisa.Visible = False
    '    End Select

    Call gM46V999.gpFecha2(lbdTab_Dizer)

End Sub

Private Sub mpCarregarComboDizer()
    cboDizer.Clear

    cboDizer.AddItem "Inclusão de veículo"
    cboDizer.ItemData(cboDizer.NewIndex) = 1

    cboDizer.AddItem "Substituição do veículo segurado"
    cboDizer.ItemData(cboDizer.NewIndex) = 2

    cboDizer.AddItem "Exclusão do veículo segurado"
    cboDizer.ItemData(cboDizer.NewIndex) = 4

    cboDizer.AddItem "Alteração das características do veículo segurado"
    cboDizer.ItemData(cboDizer.NewIndex) = 6

    cboDizer.AddItem "Alteração da importância(s) segurada(s) do veículo segurado"
    cboDizer.ItemData(cboDizer.NewIndex) = 7

    cboDizer.AddItem "Alteração do(s) acessório(s)"
    cboDizer.ItemData(cboDizer.NewIndex) = 8

    cboDizer.AddItem "Inclusão do(s) acessório(s)"
    cboDizer.ItemData(cboDizer.NewIndex) = 9

    cboDizer.AddItem "Cancelamento da apólice"
    cboDizer.ItemData(cboDizer.NewIndex) = 10

    cboDizer.AddItem "Outros"
    cboDizer.ItemData(cboDizer.NewIndex) = 12

    cboDizer.AddItem "Alteração da importância segurada RCF   "
    cboDizer.ItemData(cboDizer.NewIndex) = 13

    cboDizer.AddItem "Inclusão da cobertura de RCF"
    cboDizer.ItemData(cboDizer.NewIndex) = 14

    cboDizer.AddItem "Exclusão da franquia facultativa"
    cboDizer.ItemData(cboDizer.NewIndex) = 19

    cboDizer.AddItem "Exclusão da cobertura de RCF"
    cboDizer.ItemData(cboDizer.NewIndex) = 21

    cboDizer.AddItem "Inclusão de franquia"
    cboDizer.ItemData(cboDizer.NewIndex) = 22

    cboDizer.AddItem "Exclusão do(s) acessório(s)"
    cboDizer.ItemData(cboDizer.NewIndex) = 23

    cboDizer.AddItem "Inclusão da cobertura de APP"
    cboDizer.ItemData(cboDizer.NewIndex) = 24

    cboDizer.AddItem "Exclusão da cobertura de APP"
    cboDizer.ItemData(cboDizer.NewIndex) = 25

    cboDizer.AddItem "Inclusão de Assist.24 horas e/ou Carro reserva"
    cboDizer.ItemData(cboDizer.NewIndex) = 26

    cboDizer.ListIndex = -1


End Sub
Public Sub mpInstanciarForm(pNossoNumero As String)
    mNossoNumero = pNossoNumero
    Me.Show vbModal
    pNossoNumero = mNossoNumero
End Sub

Private Function mfEndosso() As Boolean
'Função: processa endosso

    Dim lMensagem As String    'Mensagem.

    mfEndosso = False

    '1. Consiste número de apólice anterior.
    If Not gM46V999.gfPreenchido(txtApolice.Text) Then
        Call gM46V999.gpGraLog(1, "FGE0069 - Número da apólice inválido (não preenchido).")
        If txtApolice.Visible = True And txtApolice.Enabled = True Then
            txtApolice.SetFocus
        End If
        Exit Function
    End If
    If Not IsNumeric(txtApolice.Text) Then
        Call gM46V999.gpGraLog(1, "FGE0069 - Número da apólice inválido (não numérico).")
        If txtApolice.Visible = True And txtApolice.Enabled = True Then
            txtApolice.SetFocus
        End If
        Exit Function
    End If
    If Len(Trim(txtApolice.Text)) < 10 Then
        txtApolice.Text = String(10 - Len(Trim$(txtApolice.Text)), "0") & Trim$(txtApolice.Text)
    End If

    '2. Consiste número do item da apólice.
    If optItem(1) = True Then   'Apenas um item.
        If Not gM46V999.gfPreenchido(txtItem.Text) Then
            Call gM46V999.gpGraLog(1, "FGE0070 - Número do item da apólice inválido (não preenchido).")
            If txtItem.Visible = True And txtItem.Enabled = True Then
                txtItem.SetFocus
            End If
            Exit Function
        End If
        If Not IsNumeric(txtItem.Text) Then
            Call gM46V999.gpGraLog(1, "FGE0070 - Número do item da apólice inválido (não numérico).")
            If txtItem.Visible = True And txtItem.Enabled = True Then
                txtItem.SetFocus
            End If
            Exit Function
        End If
        If Len(Trim(txtItem.Text)) < 4 Then
            txtItem.Text = Format(Trim$(txtItem.Text), "0000")
            txtItem.Text = Trim$(txtItem.Text)
        End If
    End If

    '3. Grava registros de cotação/proposta para renovação Yasuda.
    If mfGraEndosso = False Then
        If gM46V999.gfPreenchido(lMensagem) Then
            Call gM46V999.gpGraLog(1, lMensagem)
        End If
        If txtApolice.Visible = True And txtApolice.Enabled = True Then
            txtApolice.SetFocus
        End If
        Exit Function
    End If

    mfEndosso = True
End Function



Private Sub cmdCancelar_Click()
    gM46V999.gFormCancelado = True
    Unload Me
End Sub

Private Sub cmdOk_Click()
    If cboDizer.ListIndex = -1 Then
        Call gM46V999.gpGraLog(2, "FGE0114 - Tipo de endosso não selecionado.")
        If cboDizer.Visible = True And cboDizer.Enabled = True Then
            cboDizer.SetFocus
        End If
        Exit Sub
        '    ElseIf cboDizer.ItemData(cboDizer.ListIndex) = 12 Then  'Outros
        '        If gM46V999.gfPreenchido(txtDizer.Text) = False Then
        '            Call gM46V999.gpGraLog(2, "FNP0046 - Descrição do endosso não preenchido.")
        '            If txtDizer.Visible = True And txtDizer.Enabled = True Then
        '                txtDizer.SetFocus
        '            End If
        '            Exit Sub
        '        Else
        '            If Len(Trim(txtDizer.Text)) < 15 Then
        '                Call gM46V999.gpGraLog(2, "FGE0114 - Descrição do endosso inválida (Menos de 15 caracteres).")
        '                If txtDizer.Visible = True And txtDizer.Enabled = True Then
        '                    txtDizer.SetFocus
        '                End If
        '                Exit Sub
        '            End If
        '        End If
    End If

    If mCod_Ramo = 112 And _
       (mCod_Modalidade = 1020 Or _
        mCod_Modalidade = 1030 Or _
        mCod_Modalidade = 1040 Or _
        mCod_Modalidade = 1050 Or _
        mCod_Modalidade = 1060) Then
        If cboDizer.ItemData(cboDizer.ListIndex) = 5 Or _
           cboDizer.ItemData(cboDizer.ListIndex) = 7 Then
            Call gM46V999.gpGraLog(2, "FGE0114 - Não são permitidos esses dizer para essa apólice.")
            If txtDizer.Visible = True And txtDizer.Enabled = True Then
                txtDizer.SetFocus
            End If
            Exit Sub
        End If
    End If

    If mCod_Ramo = 113 Then
        If Not (cboDizer.ItemData(cboDizer.ListIndex) = 3 Or _
                cboDizer.ItemData(cboDizer.ListIndex) = 6 Or _
                cboDizer.ItemData(cboDizer.ListIndex) = 10 Or _
                cboDizer.ItemData(cboDizer.ListIndex) = 11) Then
            Call gM46V999.gpGraLog(2, "FGE0114 - Não são permitidos esses dizer para esse ramo.")
            If txtDizer.Visible = True And txtDizer.Enabled = True Then
                txtDizer.SetFocus
            End If
            Exit Sub
        End If
    End If

    Screen.MousePointer = vbHourglass
    If mfEndosso = False Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    txtApolice.Text = ""
    txtItem.Text = ""
    Me.Hide
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
    frmT46V001A.panForm = Mid$(Me.Name, 4)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmT46V011A = Nothing
End Sub


Private Sub lstDizer_Click()

    If lstDizer.ListIndex <> -1 Then
        txtDizer.Text = mArrayDizerOutros(lstDizer.ListIndex)
    Else
        txtDizer.Text = ""
    End If
    '    If lstDizer.ListIndex = 0 And gM46V999.PerfilAcesso = e_Perfil_DIG_IST Then
    '        txtDizer.Enabled = False
    '    Else
    '        txtDizer.Enabled = True
    '    End If


End Sub

Private Sub optItem_Click(Index As Integer)
    Select Case Index
    Case 0      'Todos os itens.
        txtItem.Text = ""
        txtItem.Enabled = False
        txtItem.BackColor = &H8000000F
    Case 1      'Apenas o item.
        txtItem.Enabled = True
        txtItem.BackColor = vbWhite
    End Select
End Sub

Private Sub optItem_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub txtApolice_GotFocus()
    With txtApolice
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtApolice_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtDizer_GotFocus()
    With txtDizer
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtDizer_LostFocus()
'txtDizer.Text = UCase(gpRetirarCaracterEspecial(txtDizer.Text))
End Sub

Private Sub txtItem_GotFocus()
    With txtItem
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub mpAjustaForm()
'Procedure utilizada para posicionar os objetos deste formulario
    Me.Height = 2820
    cmdCancelar.Top = 2040
    cmdOk.Top = 2040
    panEndosso.Height = 1740
    txtApolice.Top = 1080
    cmdPesquisa.Top = 1080
    lblX(0).Top = 1080
    optItem(0).Top = 960
    optItem(1).Top = 1320
    txtItem.Top = 1320
    txtDizer.Visible = False
    fraDizer.Visible = False
End Sub

