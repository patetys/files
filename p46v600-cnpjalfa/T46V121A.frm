VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT46V121A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ajustes de data de parcelamento"
   ClientHeight    =   5535
   ClientLeft      =   4440
   ClientTop       =   3450
   ClientWidth     =   2940
   ControlBox      =   0   'False
   Icon            =   "T46V121A.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5535
   ScaleWidth      =   2940
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtDtVenc 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   12
      Left            =   1335
      TabIndex        =   35
      ToolTipText     =   "Data de vencimento da décima-primeira parcela."
      Top             =   4560
      Width           =   1000
   End
   Begin VB.TextBox txtDtVenc 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   11
      Left            =   1320
      TabIndex        =   10
      ToolTipText     =   "Data de vencimento da décima-primeira parcela."
      Top             =   4200
      Width           =   1000
   End
   Begin VB.TextBox txtDtVenc 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   10
      Left            =   1320
      TabIndex        =   9
      ToolTipText     =   "Data de vencimento da décima parcela."
      Top             =   3840
      Width           =   1000
   End
   Begin VB.TextBox txtDtVenc 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   9
      Left            =   1320
      TabIndex        =   8
      ToolTipText     =   "Data de vencimento da nona parcela."
      Top             =   3480
      Width           =   1000
   End
   Begin VB.TextBox txtDtVenc 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   8
      Left            =   1320
      TabIndex        =   7
      ToolTipText     =   "Data de vencimento da oitava parcela."
      Top             =   3120
      Width           =   1000
   End
   Begin VB.TextBox txtDtVenc 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   7
      Left            =   1320
      TabIndex        =   6
      ToolTipText     =   "Data de vencimento da sétima parcela."
      Top             =   2760
      Width           =   1000
   End
   Begin VB.TextBox txtDtVenc 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   6
      Left            =   1320
      TabIndex        =   5
      ToolTipText     =   "Data de vencimento da sexta parcela."
      Top             =   2400
      Width           =   1000
   End
   Begin VB.TextBox txtDtVenc 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   5
      Left            =   1320
      TabIndex        =   4
      ToolTipText     =   "Data de vencimento da quinta parcela."
      Top             =   2040
      Width           =   1000
   End
   Begin VB.TextBox txtDtVenc 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   4
      Left            =   1320
      TabIndex        =   3
      ToolTipText     =   "Data de vencimento da quarta parcela."
      Top             =   1680
      Width           =   1000
   End
   Begin VB.TextBox txtDtVenc 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   3
      Left            =   1320
      TabIndex        =   2
      ToolTipText     =   "Data de vencimento da terceira parcela."
      Top             =   1320
      Width           =   1000
   End
   Begin VB.TextBox txtDtVenc 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   2
      Left            =   1320
      TabIndex        =   1
      ToolTipText     =   "Data de vencimento da segunda parcela."
      Top             =   960
      Width           =   1000
   End
   Begin VB.TextBox txtDtVenc 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   0
      ToolTipText     =   "Data de vencimento da primeira parcela."
      Top             =   600
      Width           =   1000
   End
   Begin Threed.SSCommand cmdOk 
      Cancel          =   -1  'True
      Height          =   300
      Left            =   1755
      TabIndex        =   11
      ToolTipText     =   "Volta à proposta."
      Top             =   4995
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Ok"
      ForeColor       =   8388608
   End
   Begin Threed.SSCommand cmdVenc 
      Height          =   375
      Index           =   1
      Left            =   2400
      TabIndex        =   24
      TabStop         =   0   'False
      ToolTipText     =   "calcula as próximas datas."
      Top             =   600
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   78
      ForeColor       =   12582912
      MouseIcon       =   "T46V121A.frx":000C
      Picture         =   "T46V121A.frx":0028
   End
   Begin Threed.SSCommand cmdVenc 
      Height          =   375
      Index           =   2
      Left            =   2400
      TabIndex        =   25
      TabStop         =   0   'False
      ToolTipText     =   "calcula as próximas datas."
      Top             =   960
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   78
      ForeColor       =   12582912
      MouseIcon       =   "T46V121A.frx":0182
      Picture         =   "T46V121A.frx":019E
   End
   Begin Threed.SSCommand cmdVenc 
      Height          =   375
      Index           =   3
      Left            =   2400
      TabIndex        =   26
      TabStop         =   0   'False
      ToolTipText     =   "calcula as próximas datas."
      Top             =   1320
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   78
      ForeColor       =   12582912
      MouseIcon       =   "T46V121A.frx":02F8
      Picture         =   "T46V121A.frx":0314
   End
   Begin Threed.SSCommand cmdVenc 
      Height          =   375
      Index           =   4
      Left            =   2400
      TabIndex        =   27
      TabStop         =   0   'False
      ToolTipText     =   "calcula as próximas datas."
      Top             =   1680
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   78
      ForeColor       =   12582912
      MouseIcon       =   "T46V121A.frx":046E
      Picture         =   "T46V121A.frx":048A
   End
   Begin Threed.SSCommand cmdVenc 
      Height          =   375
      Index           =   5
      Left            =   2400
      TabIndex        =   28
      TabStop         =   0   'False
      ToolTipText     =   "calcula as próximas datas."
      Top             =   2040
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   78
      ForeColor       =   12582912
      MouseIcon       =   "T46V121A.frx":05E4
      Picture         =   "T46V121A.frx":0600
   End
   Begin Threed.SSCommand cmdVenc 
      Height          =   375
      Index           =   6
      Left            =   2400
      TabIndex        =   29
      TabStop         =   0   'False
      ToolTipText     =   "calcula as próximas datas."
      Top             =   2400
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   78
      ForeColor       =   12582912
      MouseIcon       =   "T46V121A.frx":075A
      Picture         =   "T46V121A.frx":0776
   End
   Begin Threed.SSCommand cmdVenc 
      Height          =   375
      Index           =   7
      Left            =   2400
      TabIndex        =   30
      TabStop         =   0   'False
      ToolTipText     =   "calcula as próximas datas."
      Top             =   2760
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   78
      ForeColor       =   12582912
      MouseIcon       =   "T46V121A.frx":08D0
      Picture         =   "T46V121A.frx":08EC
   End
   Begin Threed.SSCommand cmdVenc 
      Height          =   375
      Index           =   8
      Left            =   2400
      TabIndex        =   31
      TabStop         =   0   'False
      ToolTipText     =   "calcula as próximas datas."
      Top             =   3120
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   78
      ForeColor       =   12582912
      MouseIcon       =   "T46V121A.frx":0A46
      Picture         =   "T46V121A.frx":0A62
   End
   Begin Threed.SSCommand cmdVenc 
      Height          =   375
      Index           =   9
      Left            =   2400
      TabIndex        =   32
      TabStop         =   0   'False
      ToolTipText     =   "calcula as próximas datas."
      Top             =   3480
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   78
      ForeColor       =   12582912
      MouseIcon       =   "T46V121A.frx":0BBC
      Picture         =   "T46V121A.frx":0BD8
   End
   Begin Threed.SSCommand cmdVenc 
      Height          =   375
      Index           =   10
      Left            =   2400
      TabIndex        =   33
      TabStop         =   0   'False
      ToolTipText     =   "calcula as próximas datas."
      Top             =   3840
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   78
      ForeColor       =   12582912
      MouseIcon       =   "T46V121A.frx":0D32
      Picture         =   "T46V121A.frx":0D4E
   End
   Begin Threed.SSCommand cmdVenc 
      Height          =   375
      Index           =   11
      Left            =   2400
      TabIndex        =   34
      TabStop         =   0   'False
      ToolTipText     =   "calcula as próximas datas."
      Top             =   4200
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   78
      ForeColor       =   12582912
      MouseIcon       =   "T46V121A.frx":0EA8
      Picture         =   "T46V121A.frx":0EC4
   End
   Begin Threed.SSCommand cmdVenc 
      Height          =   375
      Index           =   12
      Left            =   2400
      TabIndex        =   36
      TabStop         =   0   'False
      ToolTipText     =   "calcula as próximas datas."
      Top             =   4560
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   78
      ForeColor       =   12582912
      MouseIcon       =   "T46V121A.frx":101E
      Picture         =   "T46V121A.frx":103A
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "12 parcelas"
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   12
      Left            =   255
      TabIndex        =   37
      Top             =   4560
      Width           =   1005
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1 parcela"
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   23
      Top             =   600
      Width           =   1005
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Data vencimento"
      ForeColor       =   &H00800000&
      Height          =   330
      Index           =   0
      Left            =   240
      TabIndex        =   22
      Top             =   120
      Width           =   2580
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2 parcelas"
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   21
      Top             =   930
      Width           =   1005
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "3 parcelas"
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   20
      Top             =   1320
      Width           =   1005
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "4 parcelas"
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   19
      Top             =   1680
      Width           =   1005
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "5 parcelas"
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   18
      Top             =   2040
      Width           =   1005
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "6 parcelas"
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   17
      Top             =   2400
      Width           =   1005
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "7 parcelas"
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   7
      Left            =   240
      TabIndex        =   16
      Top             =   2760
      Width           =   1005
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "8 parcelas"
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   8
      Left            =   240
      TabIndex        =   15
      Top             =   3120
      Width           =   1005
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "9 parcelas"
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   9
      Left            =   240
      TabIndex        =   14
      Top             =   3480
      Width           =   1005
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "10 parcelas"
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   10
      Left            =   240
      TabIndex        =   13
      Top             =   3840
      Width           =   1005
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "11 parcelas"
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   11
      Left            =   240
      TabIndex        =   12
      Top             =   4200
      Width           =   1005
   End
End
Attribute VB_Name = "frmT46V121A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mQtdeParcelas As Integer
Dim mPgtAntec As Boolean
Private Sub cmdOk_Click()
    Dim lCount As Integer    'utilizado no for...next
    Dim lMensagem As String     'Mensagem

    If gProChamado = "T46V012" Then
        For lCount = 1 To (mQtdeParcelas - 1)
            If lCount < 12 Then    'thais
                If gM46V999.gfPreenchido(txtDtVenc(lCount).Text) = True And txtDtVenc(lCount).Locked = False Then
                    GoTo VerData
                End If
            End If
        Next
        If txtDtVenc(1).Locked = True Then
            txtDtVenc(1).Text = ""
        End If
    End If
    If gProChamado = "T46V012A" Then
VerData:
        For lCount = 1 To (mQtdeParcelas - 1)
            If lCount < 12 Then
                If IsDate(txtDtVenc(lCount).Text) = False Then
                    Call gM46V999.gpGraLog(1, "FGE0082 - Data de vencimento da parcela " & lCount & " inválida (não preenchida).")
                    txtDtVenc(lCount).SetFocus
                    Exit Sub
                End If
            End If
        Next
    End If
    
    'Não permitir ajuste de parcelas com vencimento menor que 30 dias do termino de vigência
    If frmT46V101A.objstcPedido.Cod_Produto >= 1102 Then
        For lCount = 1 To 12
            If lCount < 12 Then
                If IsDate(txtDtVenc(lCount).Text) Then
                    If DateDiff("D", txtDtVenc(lCount).Text, gM46V999.gfForData(2, frmT46V101A.objstcPedido.Dat_Ter_Vig)) < 30 Then
                        Call gM46V999.gpGraLog(1, "FGE0121 - Não é permitido vencimentos posteriores à 30 dias do termino de vigência.")
                        txtDtVenc(lCount).SetFocus
                        Exit Sub
                    End If
                End If
            End If
        Next lCount
    End If

    Me.Hide
End Sub
Public Sub gpInstanciarForm(pQtdeParcelas As Integer, pPgtAntec As Boolean, pColecaoParc As Collection)
    Dim lI As Integer
    Dim lobjstc As stcA46V713B


    mQtdeParcelas = pQtdeParcelas
    mPgtAntec = pPgtAntec
    If pPgtAntec = True And frmT46V101A.objstcPedido.Cod_Estip <> 20 Then 'Não aplicar para o imobiliário
        txtDtVenc(1).Text = Format(Date + 8, "dd/mm/yyyy")
    End If

    For Each lobjstc In pColecaoParc
        txtDtVenc(lobjstc.Num_Parc).Text = gM46V999.gfForData(2, lobjstc.DAT_VENC)
    Next lobjstc
    Me.Show vbModal

End Sub

Private Sub cmdVenc_Click(Index As Integer)
    Dim lI As Integer
    txtDtVenc(Index).Text = gM46V999.gfForData(9, txtDtVenc(Index).Text)
    If txtDtVenc(index).Text <> "" And index < 12 Then    ' originalmente sem o conteudo a direita do operador AND
        For lI = (Index + 1) To gM46V999.gProQuaMaxParcelas    '<- originalmente 12
            If txtDtVenc(lI).Enabled = True Then
                txtDtVenc(lI).Text = Format(DateAdd("m", 1, txtDtVenc(lI - 1).Text), "dd/mm/yyyy")
            Else
                Exit For
            End If
        Next
    End If
End Sub

Private Sub Form_Load()
    Dim lI As Integer      'Utilizada no for...next
'1. Posiciona form.
    Call gM46V999.gpCenForm(Me)

    If gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Or gM46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda Or gM46V999.PerfilAcesso = e_Perfil_DIG_IST Then
        txtDtVenc(1).Locked = True
    End If
    'Ficha 901824
    If gProChamado = "T46V012A" Or gProChamado = "T46V012" Or gProChamado = "T46V121A" Then
        If mPgtAntec = True Then
            txtDtVenc(1).Locked = True
            If frmT46V101A.objstcPedido.Cod_Estip <> 20 Then
            	txtDtVenc(1).Text = Date + 8
            End If
            txtDtVenc(1).ForeColor = vbRed
        End If
    Else
        'Ficha 900046
        If frmT46V102A.optPagAntecipado(2).Value = False Then
            If frmT46V101A.objstcPedido.Cod_Estip <> 20 Then
            	txtDtVenc(1).Text = Date + 8
            End If
            txtDtVenc(1).ForeColor = vbRed
        Else
            If txtDtVenc(1).Text = Date + 8 Then
                txtDtVenc(1).Text = ""
                txtDtVenc(1).ForeColor = txtDtVenc(2).ForeColor
            End If
        End If
    End If
End Sub

Private Sub Form_Activate()
    Dim lI As Integer
    frmT46V100A.panForm.Caption = Mid$(Me.Name, 4)
    'Ficha 901824
    If gProChamado = "T46V012A" Or gProChamado = "T46V012" Or gProChamado = "T46V121A" Then
        For lI = 1 To 12
            If lI < 12 And lI <= (mQtdeParcelas) Then
                txtDtVenc(lI).Enabled = True
            Else
                txtDtVenc(lI).Enabled = False
                txtDtVenc(lI).Text = ""
            End If
        Next
        If mPgtAntec = True Then
            txtDtVenc(1).Locked = True
            If frmT46V101A.objstcPedido.Cod_Estip <> 20 Then
            	txtDtVenc(1).Text = Date + 7
            End If
            txtDtVenc(1).ForeColor = vbRed
        Else
            txtDtVenc(1).Locked = False
            If txtDtVenc(1).Text = Date + 7 Then
                'txtDtVenc(0).Text = ""
            End If
            txtDtVenc(1).ForeColor = txtDtVenc(2).ForeColor
        End If
    Else
        If Not gM46V999.gfPreenchido(frmT46V102A.txtQuaParcelas.Text) Then
            frmT46V102A.txtQuaParcelas.Text = "0"
        End If

        For lI = 1 To 12
            If lI <= gM46V999.gProQuaMaxParcelas And lI <= CInt(frmT46V102A.txtQuaParcelas.Text) Then
                txtDtVenc(lI).Enabled = True
                txtDtVenc(lI).Visible = True
                lblX(lI).Enabled = True
                lblX(lI).Visible = True
                cmdVenc(lI).Enabled = True
                cmdVenc(lI).Visible = True
            Else
                txtDtVenc(lI).Enabled = False
                txtDtVenc(lI).Text = ""
            End If
        Next

        If CInt(frmT46V102A.txtQuaParcelas.Text) < 12 Then
            Dim cont_desabilita As Byte
            For cont_desabilita = CInt(frmT46V102A.txtQuaParcelas.Text) + 1 To 12
                lblX(cont_desabilita).Enabled = False
                txtDtVenc(cont_desabilita).Enabled = False
                cmdVenc(cont_desabilita).Enabled = False
            Next cont_desabilita
        End If

        'Ficha 900046
        If frmT46V102A.optPagAntecipado(2).Value = False Then
            If gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Or gM46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda Or gM46V999.PerfilAcesso = e_Perfil_DIG_IST Then
                txtDtVenc(1).Locked = True
            End If
            If frmT46V101A.objstcPedido.Cod_Estip <> 20 Then
            	txtDtVenc(1).Text = Date + 8
            	txtDtVenc(1).ForeColor = vbRed
            End If
        Else
            txtDtVenc(1).Locked = False
            txtDtVenc(1).ForeColor = txtDtVenc(2).ForeColor
        End If

    End If
Finalizar:
    If gM46V999.gDesabilitarForms Then
        If Me.Tag <> "Desabilitado" Then
            'Ficha 900654
            If frmT46V101A.objstcPedido.mTransmissao = False Then
                DesabilitaControles Me
            End If
            cmdOk.Enabled = True
        End If
    Else
        If Me.Tag = "Desabilitado" Then
            HabilitaControles Me
        End If
    End If

End Sub

Private Sub txtDtVenc_Change(Index As Integer)
'Ficha 901824
    If gProChamado <> "T46V012A" And gProChamado <> "T46V012" And gProChamado <> "T46V121A" Then
        Call gpHabSalvar    'Posiciona controles de salvar e desfazer.
    End If
End Sub

Private Sub txtDtVenc_GotFocus(Index As Integer)
    With txtDtVenc(Index)
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtDtVenc_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtDtVenc_LostFocus(Index As Integer)
    If gCarga Then Exit Sub
    txtDtVenc(Index).Text = gM46V999.gfForData(9, txtDtVenc(Index).Text)
    'Ficha 901824
    If gProChamado <> "T46V012A" And gProChamado <> "T46V012" And gProChamado <> "T46V121A" Then
        If Index < CInt(frmT46V102A.txtQuaParcelas.Text) Then    '<> 12
            If txtDtVenc(Index + 1).Text = "" And txtDtVenc(Index + 1).Enabled = True And txtDtVenc(Index).Text <> "" Then
                txtDtVenc(Index + 1).Text = DateAdd("m", 1, txtDtVenc(Index).Text)
            End If
        End If
    End If
    If gM46V999.gfPreenchido(txtDtVenc(12).Text) Then
        gM46V999.gProDatUltParcela = gM46V999.gfForData(1, txtDtVenc(12).Text)
    End If
End Sub

Public Sub gpInstanciarForm2(pQtdeParcelas As Integer, pPgtAntec As Boolean, pArrayDatas As Variant)
    Dim lI As Integer
    Dim lobjstc As stcA46V713B

    mQtdeParcelas = pQtdeParcelas
    mPgtAntec = pPgtAntec
    If pPgtAntec = True And frmT46V101A.objstcPedido.Cod_Estip <> 20 Then
        txtDtVenc(1).Text = Format(Date + 8, "dd/mm/yyyy")
    End If
    If IsEmpty(pArrayDatas) = False Then
        For lI = 0 To UBound(pArrayDatas)
            If gM46V999.gfPreenchido(pArrayDatas(lI)) Then
                frmT46V121A.txtDtVenc(lI + 1).Text = gM46V999.gfForData(2, pArrayDatas(lI))
            End If
        Next lI
    End If

End Sub

