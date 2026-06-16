VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT47V019A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ajustes de data de parcelamento"
   ClientHeight    =   4980
   ClientLeft      =   4440
   ClientTop       =   3450
   ClientWidth     =   2940
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4980
   ScaleWidth      =   2940
   ShowInTaskbar   =   0   'False
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
      Left            =   1785
      TabIndex        =   11
      ToolTipText     =   "Volta à proposta."
      Top             =   4650
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
      MouseIcon       =   "T47V019A.frx":0000
      Picture         =   "T47V019A.frx":001C
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
      Picture         =   "T47V019A.frx":0176
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
      Picture         =   "T47V019A.frx":02D0
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
      Picture         =   "T47V019A.frx":042A
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
      Picture         =   "T47V019A.frx":0584
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
      Picture         =   "T47V019A.frx":06DE
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
      Picture         =   "T47V019A.frx":0838
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
      Picture         =   "T47V019A.frx":0992
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
      Picture         =   "T47V019A.frx":0AEC
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
      Picture         =   "T47V019A.frx":0C46
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
      Picture         =   "T47V019A.frx":0DA0
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
Attribute VB_Name = "frmT47v019A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mQtdeParcelas   As Integer
Dim mPgtAntec       As Boolean
Private Sub cmdOK_Click()
Dim lCount      As Integer      'utilizado no for...next
Dim lMensagem   As String       'Mensagem
    
    For lCount = 1 To (mQtdeParcelas - 1)
        If lCount < 11 Then
            If IsDate(txtDtVenc(lCount).Text) = False Then
                Call gM46V999.gpGraLog(1, "FGE0082 - Data de vencimento da parcela " & lCount & " inválida (não preenchida).")
                txtDtVenc(lCount).SetFocus
                Exit Sub
            End If
        End If
    Next
    
    Me.Hide
End Sub
Public Sub gpInstanciarForm(pQtdeParcelas As Integer, pPgtAntec As Boolean, pColecaoParc As Collection)
Dim lI  As Integer
Dim lobjstc As stcA46V713B


    mQtdeParcelas = pQtdeParcelas
    mPgtAntec = pPgtAntec
    If pPgtAntec = True Then
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
    If txtDtVenc(Index).Text <> "" And Index < 11 Then ' originalmente sem o conteudo a direita do operador AND
        For lI = (Index + 1) To gM46V999.gProQuaMaxParcelas '<- originalmente 12
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
    If frmT47V004A.optPagAntecipado(2).Value = False Then
        txtDtVenc(1).Text = Date + 8
        txtDtVenc(1).ForeColor = vbRed
    Else
        If txtDtVenc(1).Text = Date + 8 Then
            txtDtVenc(1).Text = ""
            txtDtVenc(1).ForeColor = txtDtVenc(2).ForeColor
        End If
    End If
End Sub

Private Sub Form_Activate()
Dim lI As Integer

    If Not gM46V999.gfPreenchido(frmT47V004A.txtQuaParcelas.Text) Then
        frmT47V004A.txtQuaParcelas.Text = "0"
    End If

    For lI = 1 To 11
        If lI <= gM46V999.gProQuaMaxParcelas And lI <= CInt(frmT47V004A.txtQuaParcelas.Text) Then
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
    
    If CInt(frmT47V004A.txtQuaParcelas.Text) < 11 Then
        Dim cont_desabilita As Byte
        For cont_desabilita = CInt(frmT47V004A.txtQuaParcelas.Text) + 1 To 11
            lblX(cont_desabilita).Enabled = False
            txtDtVenc(cont_desabilita).Enabled = False
            cmdVenc(cont_desabilita).Enabled = False
        Next cont_desabilita
    End If

'Ficha 900046
    If frmT47V004A.optPagAntecipado(2).Value = False Then
        If gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Or gM46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda Or gM46V999.PerfilAcesso = e_Perfil_DIG_IST Then
            txtDtVenc(1).Locked = True
        End If
        txtDtVenc(1).Text = Date + 8
        txtDtVenc(1).ForeColor = vbRed
    Else
        txtDtVenc(1).Locked = False
        txtDtVenc(1).ForeColor = txtDtVenc(2).ForeColor
    End If
        
Finalizar:
    If gM46V999.gDesabilitarForms Then
        If Me.Tag <> "Desabilitado" Then
            DesabilitaControles Me
        End If
    Else
        If Me.Tag = "Desabilitado" Then
            HabilitaControles Me
        End If
    End If
    
End Sub

Private Sub txtDtVenc_Change(Index As Integer)
'Ficha 901824
    If gProChamado <> "T47V019A" Then
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
    If gProChamado <> "T47V019A" Then
        If Index < CInt(frmT47V004A.txtQuaParcelas.Text) Then '<> 12
            If txtDtVenc(Index + 1).Text = "" And txtDtVenc(Index + 1).Enabled = True And txtDtVenc(Index).Text <> "" Then
                txtDtVenc(Index + 1).Text = DateAdd("m", 1, txtDtVenc(Index).Text)
            End If
        End If
    End If
    If gM46V999.gfPreenchido(txtDtVenc(11).Text) Then
        gM46V999.gProDatUltParcela = gM46V999.gfForData(1, txtDtVenc(11).Text)
    End If
End Sub

Public Sub gpInstanciarForm2(pQtdeParcelas As Integer, pPgtAntec As Boolean, pArrayDatas As Variant)
Dim lI  As Integer
Dim lobjstc As stcA46V713B

    mQtdeParcelas = pQtdeParcelas
    mPgtAntec = pPgtAntec
    If pPgtAntec = True Then
        txtDtVenc(1).Text = Format(Date + 8, "dd/mm/yyyy")
    End If
    If IsEmpty(pArrayDatas) = False Then
        For lI = 0 To UBound(pArrayDatas)
            If gM46V999.gfPreenchido(pArrayDatas(lI)) Then
                frmT47v019A.txtDtVenc(lI + 1).Text = gM46V999.gfForData(2, pArrayDatas(lI))
            End If
        Next lI
    End If

End Sub

