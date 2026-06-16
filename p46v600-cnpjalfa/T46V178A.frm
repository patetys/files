VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT46V178A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Endosso"
   ClientHeight    =   2100
   ClientLeft      =   5505
   ClientTop       =   3270
   ClientWidth     =   9795
   Icon            =   "T46V178A.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   9795
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel panEndosso 
      Height          =   1830
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9750
      _Version        =   65536
      _ExtentX        =   17198
      _ExtentY        =   3228
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
      Begin VB.Frame fmTipoEndosso 
         Caption         =   "Tipo de endosso:"
         ForeColor       =   &H00800000&
         Height          =   810
         Left            =   150
         TabIndex        =   3
         Top             =   120
         Width           =   9105
         Begin Threed.SSOption optTipoEndosso 
            Height          =   315
            Index           =   0
            Left            =   300
            TabIndex        =   6
            Top             =   300
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   556
            _StockProps     =   78
            Caption         =   "Alterações"
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
         Begin Threed.SSOption optTipoEndosso 
            Height          =   315
            Index           =   1
            Left            =   1800
            TabIndex        =   7
            Top             =   300
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   556
            _StockProps     =   78
            Caption         =   "Inclusão de item"
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
         Begin Threed.SSOption optTipoEndosso 
            Height          =   315
            Index           =   2
            Left            =   3675
            TabIndex        =   8
            Top             =   300
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   556
            _StockProps     =   78
            Caption         =   "Exclusão de item"
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
         Begin Threed.SSOption optTipoEndosso 
            Height          =   315
            Index           =   3
            Left            =   5550
            TabIndex        =   9
            Top             =   300
            Width           =   2340
            _Version        =   65536
            _ExtentX        =   4128
            _ExtentY        =   556
            _StockProps     =   78
            Caption         =   "Cancelamento da Apólice"
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
      Begin VB.TextBox txtApolice 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1305
         MaxLength       =   10
         TabIndex        =   2
         ToolTipText     =   "Número da apólice anterior."
         Top             =   1140
         Width           =   1100
      End
      Begin VB.TextBox txtItem 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   4905
         MaxLength       =   4
         TabIndex        =   1
         ToolTipText     =   "Número da cotação."
         Top             =   1140
         Width           =   600
      End
      Begin Threed.SSCommand cmdPesquisa 
         Height          =   285
         Left            =   2505
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Busca apólice."
         Top             =   1140
         Width           =   285
         _Version        =   65536
         _ExtentX        =   494
         _ExtentY        =   503
         _StockProps     =   78
         ForeColor       =   12582912
         Picture         =   "T46V178A.frx":030A
      End
      Begin Threed.SSCommand cmdOk 
         Height          =   300
         Left            =   7050
         TabIndex        =   10
         Top             =   1200
         Width           =   1005
         _Version        =   65536
         _ExtentX        =   1764
         _ExtentY        =   529
         _StockProps     =   78
         Caption         =   "&Ok"
         ForeColor       =   8388608
      End
      Begin Threed.SSCommand cmdCancelar 
         Cancel          =   -1  'True
         Height          =   300
         Left            =   8175
         TabIndex        =   11
         Top             =   1200
         Width           =   1005
         _Version        =   65536
         _ExtentX        =   1764
         _ExtentY        =   529
         _StockProps     =   78
         Caption         =   "&Cancelar"
         ForeColor       =   8388608
      End
      Begin VB.Label lblX 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Item: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   1
         Left            =   4050
         TabIndex        =   12
         Top             =   1125
         Width           =   795
      End
      Begin VB.Label lblX 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Apólice: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   450
         TabIndex        =   5
         Top             =   1140
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmT46V178A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mNossoNumero As String

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim i As Integer
    Dim AuxValida As String


    AuxValida = "0"
    For i = 0 To optTipoEndosso.Count - 1
        If optTipoEndosso(i).Value Then
            AuxValida = "1"
        End If
    Next
    If AuxValida <> "1" Then
        Call gM46V999.gpGraLog(1, "FGE0114 - Tipo de endosso não selecionado.")
        Exit Sub
    End If

    'Limpa a variavel caso esteja preenchida para não dar erro de chave na geração do novo endosso
    If Len(mNossoNumero) > 0 Then
        mNossoNumero = ""
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

Private Function mfEndosso() As Boolean
'Função: processa endosso

    Dim lMensagem As String    'Mensagem.
    Dim lTdo As Integer
    Dim lCodDizer As Integer

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
    If Val(txtApolice.Text) < 1 Then
        Call gM46V999.gpGraLog(1, "FGE0069 - Número da apólice inválido (zerado).")
        If txtApolice.Visible = True And txtApolice.Enabled = True Then
            txtApolice.SetFocus
        End If
        Exit Function
    End If
    If Len(Trim(txtApolice.Text)) < 10 Then
        txtApolice.Text = String(10 - Len(Trim$(txtApolice.Text)), "0") & Trim$(txtApolice.Text)
    End If

    '2. Consiste número do item da apólice.
    If Not (optTipoEndosso(1).Value Or optTipoEndosso(3).Value) Then     'Se não for endosso de inclusao
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
        If Val(txtItem.Text) < 1 Then
            Call gM46V999.gpGraLog(1, "FGE0070 - Número do item da apólice inválido (zerado).")
            If txtItem.Visible = True And txtItem.Enabled = True Then
                txtItem.SetFocus
            End If
            Exit Function
        End If

        If Len(Trim(txtItem.Text)) < 4 Then
            txtItem.Text = Format(Trim$(txtItem.Text), "0000")
            txtItem.Text = Trim$(txtItem.Text)
        End If
    Else
        txtItem.Text = 0
    End If

    If optTipoEndosso(0) Then
        lCodDizer = e_CodDizerAlteracao
    ElseIf optTipoEndosso(1) Then
        lCodDizer = e_CodDizerInclusao
    ElseIf optTipoEndosso(2) Then
        lCodDizer = e_CodDizerexclusao
    ElseIf optTipoEndosso(3) Then
        lCodDizer = e_CodDizerCancelamento
    End If


    '3. Grava registros de cotação/proposta para renovação Yasuda.
    If gM46V001.gfGraEndossoYasuda(txtApolice.Text, txtItem.Text, mNossoNumero, _
                                   e_SimNao_Nao, lCodDizer, lMensagem, True, Nothing, "") = False Then
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

Private Sub cmdPesquisa_Click()
    Dim lNum_Apolice As String    'Número da apólice.
    Dim lNum_Item As String    'Número do item.

    Call frmT46V003A.mpInstanciarForm(lNum_Apolice, lNum_Item)
    txtApolice.Text = lNum_Apolice
    txtItem.Text = lNum_Item
    If cmdOk.Visible = True And cmdOk.Enabled = True Then
        cmdOk.SetFocus
    End If

End Sub

Private Sub Form_Activate()
    Call gM46V999.gpCenForm(Me)
End Sub
Public Function mpInstanciarForm(pNossoNumero As String)
    mNossoNumero = pNossoNumero
    Me.Show vbModal
    pNossoNumero = mNossoNumero
End Function

