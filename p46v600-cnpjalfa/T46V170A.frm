VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmT46V170A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista de corretores com co-corretagem"
   ClientHeight    =   5325
   ClientLeft      =   2205
   ClientTop       =   975
   ClientWidth     =   6585
   Icon            =   "T46V170A.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picChecked 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000A&
      Height          =   255
      Left            =   0
      Picture         =   "T46V170A.frx":030A
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picUnchecked 
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Frame fraCorretagem 
      Caption         =   "Corretor:"
      ForeColor       =   &H00C00000&
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   4320
      Width           =   6255
      Begin VB.TextBox txtCorretor 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Primeiro corretor."
         Top             =   360
         Width           =   5460
      End
      Begin Threed.SSCommand cmdAdicionar 
         Height          =   495
         Left            =   5640
         TabIndex        =   3
         ToolTipText     =   "Adiciona os filtros selecionados."
         Top             =   240
         Width           =   495
         _Version        =   65536
         _ExtentX        =   873
         _ExtentY        =   873
         _StockProps     =   78
         Picture         =   "T46V170A.frx":084C
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flxCorretagem 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   7223
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "frmT46V170A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bdGed As ADODB.Connection

Private Sub cmdAdicionar_Click()
    Dim lSql As String
    Dim lMensagem As String

    If gM46V999.gfPreenchido(txtCorretor.Tag) Then
        'Insere no banco de dados.

        lSql = "Delete from ged.dbo.tab_co_corretagem where cod_corretor  = " & Val(txtCorretor.Tag)
        If gM46V999.gfExeSQL(bdGed, lSql, lMensagem) <> 0 Then
            Call gM46V999.gpGraLog(1, lMensagem)
            Exit Sub
        End If

        lSql = "Insert into ged.dbo.tab_co_corretagem (cod_corretor) values ( " & Val(txtCorretor.Tag) & ")"
        If gM46V999.gfExeSQL(bdGed, lSql, lMensagem) <> 0 Then
            Call gM46V999.gpGraLog(1, lMensagem)
            Exit Sub
        End If

        Call mpCarregaTabela
    End If
    txtCorretor.Text = ""
End Sub

Private Sub mpCarregaTabela()
    Dim lSelect As String
    Dim lrst As clsYasRecordSet
    Dim lMensagem As String


    If Not gM46V999.gfAbrBasDados("GED", gM46V999.gTipGed, gM46V999.gEndGed, bdGed, lMensagem) Then
        Call gM46V999.gpGraLog(1, lMensagem)
        Exit Sub
    End If

    lSelect = " Select c.cod_corretor, corretor from ged.dbo.tab_co_corretagem as cc " & _
            " inner join ged.dbo.tbs_corretor as c on cc.cod_corretor = c.cod_corretor " & _
            " order by cc.cod_corretor"

    If Not gM46V999.gfObtRegistro(bdGed, lSelect, lrst, lMensagem) Then
        Call gM46V999.gpGraLog(1, lMensagem)
        Exit Sub
    End If

    With flxCorretagem
        .Rows = 1
        .Cols = 3
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Código"
        .TextMatrix(0, 2) = "Descrição"
        .ColWidth(0) = 195
        .ColWidth(1) = 960
        .ColWidth(2) = 4530

        Do While Not lrst.EOF
            .Rows = .Rows + 1
            .Row = .Rows - 1
            .Col = 0
            Set .CellPicture = picChecked
            .TextMatrix(.Rows - 1, 1) = lrst("Cod_Corretor")
            .TextMatrix(.Rows - 1, 2) = lrst("Corretor")
            lrst.MoveNext
        Loop

    End With

End Sub

Private Sub flxCorretagem_DblClick()
    Dim lSql As String
    Dim lMensagem As String

    If flxCorretagem.RowSel > 0 Then
        If flxCorretagem.ColSel = 0 Then
            If MsgBox("Confirma a exclusão do corretor '" & _
                      flxCorretagem.TextMatrix(flxCorretagem.RowSel, 2) & _
                      "'  da lista de co-corretagem?", vbQuestion + vbYesNo) = vbYes Then
                lSql = "Delete from ged.dbo.tab_co_corretagem where cod_corretor  = " & Val(flxCorretagem.TextMatrix(flxCorretagem.RowSel, 1))
                If gM46V999.gfExeSQL(bdGed, lSql, lMensagem) <> 0 Then
                    Call gM46V999.gpGraLog(1, lMensagem)
                    Exit Sub
                End If
                Call mpCarregaTabela
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim lMensagem As String

    If Not gM46V999.gfAbrBasDados("GED", gM46V999.gTipGed, gM46V999.gEndGed, bdGed, lMensagem) Then
        Call gM46V999.gpGraLog(1, lMensagem)
        Exit Sub
    End If

    Call mpCarregaTabela
    Call gM46V999.gpCenForm(Me)
End Sub

Private Sub txtCorretor_Change()
    txtCorretor.Tag = ""
End Sub

Private Sub txtCorretor_LostFocus()
    Dim lCodProdutor As Integer    'Código do produtor.
    Dim lCodUniProdutiva As Integer    'Código da unidade produtiva.
    Dim lCorretor As String    'Corretor.
    Dim lI As Double             'Utilizado com For...Next.
    Dim lMensagem As String    'Mensagem.
    Dim lNomProdutor As String    'Nome do produtor.
    Dim lNomUniProdutiva As String    'Nome da unidade produtiva.
    Dim lQuaCorretores As Double    'Quantidade de corretores selecionados.
    Dim lSaida() As Variant    'Resultado da busca de corretores.
    Dim lTotal As Double             'Total
    Screen.MousePointer = vbHourglass

    'Se código do corretor preenchido, encerrar processo.
    If gM46V999.gfPreenchido(txtCorretor.Tag) Or Not gM46V999.gfPreenchido(txtCorretor.Text) Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    'Posiciona campo com o que foi informado pelo usuário.
    lCorretor = gM46V999.gfRetirarCaracterEspecial(Trim$(txtCorretor.Text))

    '5. Se foram digitadas até 6 posições numéricas, trata de código de corretor.
    '   Obtém corretor, unidade produtiva e produtor.
    If IsNumeric(lCorretor) And Len(CStr(Val(lCorretor))) <= 6 Then
        If gM46V999.gfObtCorretor(1, Val(lCorretor), lSaida, lMensagem) = False Then
            Call gM46V999.gpGraLog(1, lMensagem)
            txtCorretor.Text = ""
            If txtCorretor.Visible = True And txtCorretor.Enabled = True Then
                txtCorretor.SetFocus
            End If
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        'Se for corretor 1 e tarifa >= 877 verifico se corretor escolhido é permitido.
        txtCorretor.Text = lSaida(1, 3)
        txtCorretor.Tag = lSaida(1, 1)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    '5. Se preenchido com menos de 3 posições, encerra processo.
    If Len(lCorretor) < 3 Then
        Call gM46V999.gpGraLog(1, "FGE0075 - Corretor inválido (digitar pelo menos 3 posições).")
        txtCorretor.Text = ""
        If txtCorretor.Visible = True And txtCorretor.Enabled = True Then
            txtCorretor.SetFocus
        End If
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    '6. Obtém registros com informação do usuário.
    'a) Obtém lista de corretores
    If gM46V999.gfObtCorretor(2, lCorretor, lSaida, lMensagem) = False Then
        Call gM46V999.gpGraLog(1, lMensagem)
        txtCorretor.Text = ""
        If txtCorretor.Visible = True And txtCorretor.Enabled = True Then
            txtCorretor.SetFocus
        End If
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    'b) Se campo de saída estiver vazio, nada a fazer.
    If IsNull(lSaida) Then
        txtCorretor.Text = ""
        If txtCorretor.Visible = True And txtCorretor.Enabled = True Then
            txtCorretor.SetFocus
        End If
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    'c) Posiciona quantidade de corretores encontrados.
    lQuaCorretores = UBound(lSaida)
    'd) Se encontrou apenas 1 corretor, posiciona dados deste corretor e encerra processo.
    If lQuaCorretores = 1 Then
        txtCorretor.Text = lSaida(1, 3)
        txtCorretor.Tag = lSaida(1, 1)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    'e) Se encontrou mais de um corretor, abre janela com grade de corretores.
    Load frmT46V113A
    With frmT46V113A.flxCorretores
        .Visible = True
        .Rows = lQuaCorretores + 1
        For lI = 1 To lQuaCorretores
            .Row = lI
            .Col = 0        'Código do corretor.
            .Text = lSaida(lI, 1)
            .Col = 1        'Nome do corretor.
            .Text = lSaida(lI, 2)
            .Col = 2        'Nome abreviado do corretor.
            .Text = lSaida(lI, 3)
        Next lI
        .Row = 1
        .Col = 1
    End With
    Screen.MousePointer = vbDefault
    frmT46V113A.Show vbModal, Me
    txtCorretor.Text = gNomCorretor
    txtCorretor.Tag = gCodCorretor
    If txtCorretor.Visible = True And txtCorretor.Enabled = True Then
        txtCorretor.SetFocus
    End If
    Screen.MousePointer = vbDefault

End Sub
