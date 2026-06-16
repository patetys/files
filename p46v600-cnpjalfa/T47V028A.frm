VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmT47V028A 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Prolabore e Agenciamento"
   ClientHeight    =   5415
   ClientLeft      =   6525
   ClientTop       =   2985
   ClientWidth     =   9780
   Icon            =   "T47V028A.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   9780
   Begin VB.Frame fraOpcao 
      Caption         =   "Enviar fluxo para :"
      ForeColor       =   &H00800000&
      Height          =   4680
      Left            =   105
      TabIndex        =   10
      Top             =   45
      Width           =   9435
      Begin VB.CheckBox ChkTip_Pagto 
         Caption         =   "Seguir parcelamento da Apólice ?"
         ForeColor       =   &H00800000&
         Height          =   300
         Index           =   1
         Left            =   3225
         TabIndex        =   16
         Top             =   1155
         Width           =   2865
      End
      Begin VB.CheckBox ChkTip_Pagto 
         Caption         =   "Pagamento é retido na fonte ?"
         ForeColor       =   &H00800000&
         Height          =   300
         Index           =   0
         Left            =   3225
         TabIndex        =   4
         Top             =   810
         Width           =   2865
      End
      Begin VB.TextBox TxtComissao 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   2085
         TabIndex        =   3
         Top             =   810
         Width           =   915
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxProlabore 
         Height          =   1350
         Left            =   255
         TabIndex        =   6
         Top             =   1545
         Width           =   7170
         _ExtentX        =   12647
         _ExtentY        =   2381
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxAgenciamento 
         Height          =   1350
         Left            =   255
         TabIndex        =   7
         Top             =   3135
         Width           =   7170
         _ExtentX        =   12647
         _ExtentY        =   2381
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.TextBox TxtCorretor 
         Height          =   360
         Left            =   255
         TabIndex        =   0
         Top             =   360
         Width           =   5010
      End
      Begin VB.OptionButton optOpcao 
         Caption         =   "Agenciamento"
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   1
         Left            =   6480
         TabIndex        =   2
         Top             =   360
         Width           =   1470
      End
      Begin VB.OptionButton optOpcao 
         Caption         =   "Prolabori"
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   0
         Left            =   5340
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1035
      End
      Begin Threed.SSCommand cmdAdicionar 
         Height          =   480
         Left            =   7995
         TabIndex        =   5
         Top             =   300
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   847
         _StockProps     =   78
         Caption         =   "Add"
         ForeColor       =   8388608
      End
      Begin Threed.SSCommand cmdExcluir 
         Height          =   525
         Left            =   7470
         TabIndex        =   13
         Top             =   2025
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   926
         _StockProps     =   78
         Caption         =   "&Excluir Prolabore"
         ForeColor       =   8388608
      End
      Begin Threed.SSCommand cmdExcluir1 
         Height          =   525
         Left            =   7515
         TabIndex        =   15
         Top             =   3495
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   926
         _StockProps     =   78
         Caption         =   "&Excluir Agenci."
         ForeColor       =   8388608
      End
      Begin VB.Label Label3 
         Caption         =   "Percentual de Comissão:"
         ForeColor       =   &H00800000&
         Height          =   270
         Left            =   255
         TabIndex        =   14
         Top             =   810
         Width           =   1800
      End
      Begin VB.Label Label2 
         Caption         =   "Agenciamento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   255
         TabIndex        =   12
         Top             =   2880
         Width           =   1200
      End
      Begin VB.Label Label1 
         Caption         =   "Prolabore"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   270
         Left            =   255
         TabIndex        =   11
         Top             =   1305
         Width           =   1035
      End
   End
   Begin Threed.SSCommand cmdOk 
      Height          =   480
      Left            =   3525
      TabIndex        =   8
      Top             =   4860
      Width           =   1035
      _Version        =   65536
      _ExtentX        =   1826
      _ExtentY        =   847
      _StockProps     =   78
      Caption         =   "&Ok"
      ForeColor       =   8388608
   End
   Begin Threed.SSCommand cmdFechar 
      Height          =   480
      Left            =   4890
      TabIndex        =   9
      Top             =   4860
      Width           =   1035
      _Version        =   65536
      _ExtentX        =   1826
      _ExtentY        =   847
      _StockProps     =   78
      Caption         =   "&Fechar"
      ForeColor       =   8388608
   End
End
Attribute VB_Name = "frmT47V028A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Sub ExcluirItemGrid(GridExcluir As Integer)
Dim i As Integer
        If GridExcluir = 0 Then
            With flxProlabore
                If .Rows > 2 Then
                    .RemoveItem .Row
                Else
Limpa1:
                    For i = 0 To .Cols - 1
                        .TextMatrix(.Row, i) = ""
                    Next i
                End If
            End With
        Else
            With flxAgenciamento
                If .Rows > 2 Then
                   .RemoveItem .Row
                Else
Limpa2:
                    For i = 0 To .Cols - 1
                        .TextMatrix(.Row, i) = ""
                    Next i
                End If
            End With
        End If

End Sub

Private Sub cmdAdicionar_Click()
Dim i As Integer
Dim Corretor As String
Dim Tipo As String

    If TxtCorretor.Tag <> 0 Then
        If CDblx(txtComissao.Text) > 99 Then
            MsgBox "Valor de Comissão não permitido!"
            If txtComissao.Visible = True And txtComissao.Enabled = True Then
                txtComissao.SetFocus
            End If
            Exit Sub
        End If
        If CDblx(txtComissao.Text) = 0 Then
            MsgBox "Valor de Comissão não permitido!"
            If txtComissao.Visible = True And txtComissao.Enabled = True Then
                txtComissao.SetFocus
            End If
            Exit Sub
        End If
        Corretor = Format(TxtCorretor.Tag, String(7, "0")) & " - " & Trim(TxtCorretor.Text)
        Tipo = IIf(ChkTip_Pagto(0).Value = 1, "Retido na Fonte", "Após Fatura")
        If optOpcao(0).Value = True Then
            flxProlabore.Rows = flxProlabore.Rows + 1
            i = flxProlabore.Rows - 1
            flxProlabore.TextMatrix(i, 0) = Corretor
            flxProlabore.TextMatrix(i, 1) = Format(txtComissao.Text, "00.00")
            flxProlabore.TextMatrix(i, 2) = Tipo
        Else
            flxAgenciamento.Rows = flxAgenciamento.Rows + 1
            i = flxAgenciamento.Rows - 1
            flxAgenciamento.TextMatrix(i, 0) = Corretor
            flxAgenciamento.TextMatrix(i, 1) = Format(txtComissao.Text, "00.00")
            flxAgenciamento.TextMatrix(i, 2) = Tipo
        End If
        TxtCorretor.Text = ""
        TxtCorretor.Tag = 0
        txtComissao.Text = "0,00"
        On Error Resume Next
        flxAgenciamento.FixedRows = 1
        flxProlabore.FixedRows = 1
        If TxtCorretor.Visible = True And TxtCorretor.Enabled = True Then
            TxtCorretor.SetFocus
        End If
    End If
End Sub


Private Sub cmdCancelar_Click()

End Sub

Private Sub cmdExcluir_Click()
    Call ExcluirItemGrid(0)
End Sub


Private Sub cmdExcluir1_Click()
    Call ExcluirItemGrid(1)
End Sub


Private Sub cmdFechar_Click()
    Me.Hide
    Exit Sub
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
    Call gM46V999.gpCenForm(Me)
    
    With flxProlabore
       .Rows = 1
       .Cols = 3
       .FixedCols = 0
       .TextMatrix(0, 0) = "Corretor "
       .TextMatrix(0, 1) = "Percentual"
       .TextMatrix(0, 2) = "Tipo Pagamento"
       .ColWidth(0) = 3000
       .ColWidth(1) = 1000
       .ColWidth(2) = 2000
        .ColAlignment(0) = 1
    End With
    

    With flxAgenciamento
       .Rows = 1
       .Cols = 3
       .FixedCols = 0
       .TextMatrix(0, 0) = "Corretor "
       .TextMatrix(0, 1) = "Percentual"
       .TextMatrix(0, 2) = "Tipo Pagamento"
       .ColWidth(0) = 3000
       .ColWidth(1) = 1000
       .ColWidth(2) = 2000
       .ColAlignment(0) = 1
    End With
    
    txtComissao.Text = "0,00"
          
End Sub


Private Sub optOpcao_Click(Index As Integer)
      TxtCorretor.Text = ""
      TxtCorretor.Tag = 0
End Sub

Private Sub TxtComissao_LostFocus()
    txtComissao.Text = CDblx(txtComissao.Text)
    txtComissao.Text = Format(txtComissao.Text, "0.00")
End Sub


Private Sub txtCorretor_LostFocus()
    Dim lCorretor           As String   'Corretor.
    Dim lI                  As Double   'Utilizado com For...Next.
    Dim lMensagem           As String   'Mensagem.
    Dim lQuaCorretores      As Double   'Quantidade de corretores selecionados.
    Dim lSaida()            As Variant  'Resultado da busca de corretores.
    Dim lTotal              As Double   'Total
    Dim lTipCorretor        As String
    Screen.MousePointer = vbHourglass

  
    '3. Se nome do corretor não preenchido, encerra processo.
    If Not gM46V999.gfPreenchido(TxtCorretor.Text) Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    '4. Posiciona campo com o que foi informado pelo usuário.
    lCorretor = gM46V999.gfRetirarCaracterEspecial(Trim$(TxtCorretor.Text))

    '5. Se foram digitadas até 6 posições numéricas, trata de código de corretor.
    '   Obtém corretor, unidade produtiva e produtor.
    
    If optOpcao(0).Value = True Then lTipCorretor = "P"
    If optOpcao(1).Value = True Then lTipCorretor = "A"
    
    If IsNumeric(lCorretor) And Len(CStr(Val(lCorretor))) <= 6 Then
        If gM46V999.gfObtCorr_Pro_Agenc(lTipCorretor, 1, Val(lCorretor), lSaida, lMensagem) = False Then
            Call gM46V999.gpGraLog(1, lMensagem)
            TxtCorretor.Text = ""
            If TxtCorretor.Visible = True And TxtCorretor.Enabled = True Then
                TxtCorretor.SetFocus
            End If
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        TxtCorretor.Text = lSaida(1, 3)
        TxtCorretor.Tag = lSaida(1, 1)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    '5. Se preenchido com menos de 3 posições, encerra processo.
    If Len(lCorretor) < 3 Then
        Call gM46V999.gpGraLog(1, "FGE0075 - Corretor inválido (digitar pelo menos 3 posições).")
        TxtCorretor.Text = ""
        If TxtCorretor.Visible = True And TxtCorretor.Enabled = True Then
            TxtCorretor.SetFocus
        End If
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    '6. Obtém registros com informação do usuário.
    'a) Obtém lista de corretores
    If gM46V999.gfObtCorr_Pro_Agenc(lTipCorretor, 2, lCorretor, lSaida, lMensagem) = False Then
        Call gM46V999.gpGraLog(1, lMensagem)
        TxtCorretor.Text = ""
        If TxtCorretor.Visible = True And TxtCorretor.Enabled = True Then
            TxtCorretor.SetFocus
        End If
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    'b) Se campo de saída estiver vazio, nada a fazer.
    If IsNull(lSaida) Then
        TxtCorretor.Text = ""
        
        If TxtCorretor.Visible = True And TxtCorretor.Enabled = True Then
            TxtCorretor.SetFocus
        End If
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    'c) Posiciona quantidade de corretores encontrados.
    lQuaCorretores = UBound(lSaida)
    'd) Se encontrou apenas 1 corretor, posiciona dados deste corretor e encerra processo.
    If lQuaCorretores = 1 Then
        
        TxtCorretor.Text = lSaida(1, 3)
        TxtCorretor.Tag = lSaida(1, 1)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    'e) Se encontrou mais de um corretor, abre janela com grade de corretores.
    Load frmT47V013A
    frmT47V013A.Tag = 1
    With frmT47V013A.flxCorretores
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
    frmT47V013A.Show vbModal, Me
    TxtCorretor.Text = gNomCorrProlabAgenc
    TxtCorretor.Tag = gCodCorrProlabAgenc
    'If TxtCorretor.Visible = True And TxtCorretor.Enabled = True Then
    '    TxtCorretor.SetFocus
    'End If
          
    Screen.MousePointer = vbDefault
    
End Sub


