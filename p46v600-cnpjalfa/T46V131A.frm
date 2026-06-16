VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT46V131A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Outros seguros"
   ClientHeight    =   2670
   ClientLeft      =   1575
   ClientTop       =   4260
   ClientWidth     =   6585
   Icon            =   "T46V131A.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel panRenCongenere 
      Height          =   2100
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6540
      _Version        =   65536
      _ExtentX        =   11536
      _ExtentY        =   3704
      _StockProps     =   15
      BackColor       =   13160660
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
      Begin VB.TextBox txtConItem 
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   1
         Left            =   5640
         MaxLength       =   5
         TabIndex        =   4
         ToolTipText     =   "Item de apólice para confirmação de bônus de congênere."
         Top             =   480
         Width           =   800
      End
      Begin VB.TextBox txtConApolice 
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   1
         Left            =   4200
         MaxLength       =   15
         TabIndex        =   3
         ToolTipText     =   "Número de apólice para confirmação de bônus de congênere."
         Top             =   480
         Width           =   1440
      End
      Begin VB.ComboBox cboConSeguradora 
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   5
         ItemData        =   "T46V131A.frx":000C
         Left            =   360
         List            =   "T46V131A.frx":000E
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   18
         ToolTipText     =   "Seguradora para confirmação do bônus."
         Top             =   1680
         Width           =   3855
      End
      Begin VB.ComboBox cboConSeguradora 
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   4
         ItemData        =   "T46V131A.frx":0010
         Left            =   360
         List            =   "T46V131A.frx":0012
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   14
         ToolTipText     =   "Seguradora para confirmação do bônus."
         Top             =   1380
         Width           =   3855
      End
      Begin VB.ComboBox cboConSeguradora 
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   3
         ItemData        =   "T46V131A.frx":0014
         Left            =   360
         List            =   "T46V131A.frx":0016
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         ToolTipText     =   "Seguradora para confirmação do bônus."
         Top             =   1080
         Width           =   3855
      End
      Begin VB.ComboBox cboConSeguradora 
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   2
         ItemData        =   "T46V131A.frx":0018
         Left            =   360
         List            =   "T46V131A.frx":001A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Seguradora para confirmação do bônus."
         Top             =   780
         Width           =   3855
      End
      Begin VB.ComboBox cboConSeguradora 
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   1
         ItemData        =   "T46V131A.frx":001C
         Left            =   360
         List            =   "T46V131A.frx":001E
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Seguradora para confirmação do bônus."
         Top             =   480
         Width           =   3855
      End
      Begin VB.TextBox txtConApolice 
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   2
         Left            =   4200
         MaxLength       =   15
         TabIndex        =   7
         ToolTipText     =   "Número de apólice para confirmação de bônus de congênere."
         Top             =   780
         Width           =   1440
      End
      Begin VB.TextBox txtConApolice 
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   3
         Left            =   4200
         MaxLength       =   15
         TabIndex        =   11
         ToolTipText     =   "Número de apólice para confirmação de bônus de congênere."
         Top             =   1080
         Width           =   1440
      End
      Begin VB.TextBox txtConApolice 
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   4
         Left            =   4200
         MaxLength       =   15
         TabIndex        =   15
         ToolTipText     =   "Número de apólice para confirmação de bônus de congênere."
         Top             =   1380
         Width           =   1440
      End
      Begin VB.TextBox txtConApolice 
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   5
         Left            =   4200
         MaxLength       =   15
         TabIndex        =   19
         ToolTipText     =   "Número de apólice para confirmação de bônus de congênere."
         Top             =   1680
         Width           =   1440
      End
      Begin VB.TextBox txtConItem 
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   2
         Left            =   5640
         MaxLength       =   5
         TabIndex        =   8
         ToolTipText     =   "Item de apólice para confirmação de bônus de congênere."
         Top             =   780
         Width           =   800
      End
      Begin VB.TextBox txtConItem 
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   3
         Left            =   5640
         MaxLength       =   5
         TabIndex        =   12
         ToolTipText     =   "Item de apólice para confirmação de bônus de congênere."
         Top             =   1080
         Width           =   800
      End
      Begin VB.TextBox txtConItem 
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   4
         Left            =   5640
         MaxLength       =   5
         TabIndex        =   16
         ToolTipText     =   "Item de apólice para confirmação de bônus de congênere."
         Top             =   1380
         Width           =   800
      End
      Begin VB.TextBox txtConItem 
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   5
         Left            =   5640
         MaxLength       =   5
         TabIndex        =   20
         ToolTipText     =   "Item de apólice para confirmação de bônus de congênere."
         Top             =   1680
         Width           =   800
      End
      Begin Threed.SSCheck chkCosseguro 
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Indicador de corretor."
         Top             =   540
         Width           =   195
         _Version        =   65536
         _ExtentX        =   353
         _ExtentY        =   353
         _StockProps     =   78
         ForeColor       =   0
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
      Begin Threed.SSCheck chkCosseguro 
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Indicador de corretor."
         Top             =   840
         Width           =   195
         _Version        =   65536
         _ExtentX        =   353
         _ExtentY        =   353
         _StockProps     =   78
         ForeColor       =   -2147483642
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
      Begin Threed.SSCheck chkCosseguro 
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Indicador de corretor."
         Top             =   1140
         Width           =   195
         _Version        =   65536
         _ExtentX        =   353
         _ExtentY        =   353
         _StockProps     =   78
         Caption         =   "7"
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
      Begin Threed.SSCheck chkCosseguro 
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "Indicador de corretor."
         Top             =   1440
         Width           =   195
         _Version        =   65536
         _ExtentX        =   353
         _ExtentY        =   353
         _StockProps     =   78
         ForeColor       =   0
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
      Begin Threed.SSCheck chkCosseguro 
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   17
         ToolTipText     =   "Indicador de corretor."
         Top             =   1740
         Width           =   195
         _Version        =   65536
         _ExtentX        =   353
         _ExtentY        =   353
         _StockProps     =   78
         ForeColor       =   -2147483642
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
      Begin VB.Label lblX 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Item: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   4
         Left            =   5640
         TabIndex        =   24
         Top             =   120
         Width           =   795
      End
      Begin VB.Label lblX 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Apólice: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   3
         Left            =   4200
         TabIndex        =   23
         Top             =   120
         Width           =   1440
      End
      Begin VB.Label lblX 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Seguradora"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   8
         Left            =   360
         TabIndex        =   22
         Top             =   120
         Width           =   3855
      End
   End
   Begin Threed.SSCommand cmdOk 
      Height          =   300
      Left            =   2760
      TabIndex        =   21
      ToolTipText     =   "Volta à proposta."
      Top             =   2280
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Ok"
      ForeColor       =   8388608
   End
End
Attribute VB_Name = "frmT46V131A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Dim lOpcao As Integer
    Dim lI As Integer
    Dim lCos1 As Integer
    Dim lCos2 As Integer
    Dim lCos3 As Integer
    Dim lCos4 As Integer
    Dim lCos5 As Integer

    lOpcao = 0
    For lI = 1 To 5
        If chkCosseguro(lI).Value = True Then
            lOpcao = 1
            If cboConSeguradora(lI).ListIndex = -1 Then
                Call gM46V999.gpGraLog(1, "FGE0071 - " & gM46V999.gPrefixo & "Seguradora inválida (não selecionada).")
                If cboConSeguradora(lI).Visible = True And cboConSeguradora(lI).Enabled = True Then
                    cboConSeguradora(lI).SetFocus
                End If
                Exit Sub
            End If
            If frmT46V101A.lblTipDocumento.Caption = "2" Then   'Proposta - campos obrigatórios.
                If Not gM46V999.gfPreenchido(txtConApolice(lI).Text) Then
                    Call gM46V999.gpGraLog(1, "FGE0072 - " & gM46V999.gPrefixo & _
                                              "Número da apólice inválido (renovação de congênere - não preenchido.")
                    If txtConApolice(lI).Visible = True And txtConApolice(lI).Enabled = True Then
                        txtConApolice(lI).SetFocus
                    End If
                    Exit Sub
                End If
                If Not gM46V999.gfPreenchido(txtConItem(lI).Text) Then
                    Call gM46V999.gpGraLog(1, "FGE0073 - " & gM46V999.gPrefixo & "Número do item da apólice inválido " & _
                                              "(renovação de congênere - não preenchido.")
                    If txtConItem(lI).Visible = True And txtConItem(lI).Enabled = True Then
                        txtConItem(lI).SetFocus
                    End If
                    Exit Sub
                End If
            End If
        End If
    Next lI

    'verifico se há duplicidade
    If Me.cboConSeguradora(1).ListIndex <> -1 Then
        lCos1 = cboConSeguradora(1).ItemData(cboConSeguradora(1).ListIndex)
    End If
    If Me.cboConSeguradora(2).ListIndex <> -1 Then
        lCos2 = cboConSeguradora(2).ItemData(cboConSeguradora(2).ListIndex)
    End If
    If Me.cboConSeguradora(3).ListIndex <> -1 Then
        lCos3 = cboConSeguradora(3).ItemData(cboConSeguradora(3).ListIndex)
    End If
    If Me.cboConSeguradora(4).ListIndex <> -1 Then
        lCos4 = cboConSeguradora(4).ItemData(cboConSeguradora(4).ListIndex)
    End If
    If Me.cboConSeguradora(5).ListIndex <> -1 Then
        lCos5 = cboConSeguradora(5).ItemData(cboConSeguradora(5).ListIndex)
    End If

    If lCos1 <> 0 Then
        If lCos1 = lCos2 Or lCos1 = lCos3 Or lCos1 = lCos4 Or lCos1 = lCos5 Then
            Call gM46V999.gpGraLog(1, "FGE0093 - Outro seguro inválido. Duplicidade de seguradoras.")
            Exit Sub
        End If
    End If

    If lCos2 <> 0 Then
        If lCos2 = lCos3 Or lCos2 = lCos4 Or lCos2 = lCos5 Then
            Call gM46V999.gpGraLog(1, "FGE0093 - Outro seguro inválido. Duplicidade de seguradoras.")
            Exit Sub
        End If
    End If

    If lCos3 <> 0 Then
        If lCos3 = lCos4 Or lCos3 = lCos5 Then
            Call gM46V999.gpGraLog(1, "FGE0093 - Outro seguro inválido. Duplicidade de seguradoras.")
            Exit Sub
        End If
    End If

    If lCos4 <> 0 Then
        If lCos4 = lCos5 Then
            Call gM46V999.gpGraLog(1, "FGE0093 - Outro seguro inválido. Duplicidade de seguradoras.")
            Exit Sub
        End If
    End If

    Me.Hide
End Sub

Private Sub mpExcCosseguro(ByVal pLinha As Byte)
'Procedure: excluir cosseguro (deixa linha de cosseguro em branco).

'Parâmetro...pLinha...Linha onde será excluído cosseguro.

    Call gpHabSalvar    'Posiciona controles de salvar e desfazer.
    chkCosseguro(pLinha).Value = False
    cboConSeguradora(pLinha).ListIndex = -1
    cboConSeguradora(pLinha).Enabled = False
    txtConApolice(pLinha).Enabled = False
    txtConApolice(pLinha).Text = ""
    txtConItem(pLinha).Enabled = False
    txtConItem(pLinha).Text = ""
End Sub
Private Sub mpIncCosseguro(ByVal pLinha As Byte)
'Procedure: incluir cosseguro (obtém dados do cosseguro no form T46V117A).

'Parâmetro...pLinha...Linha onde será incluído cosseguro.

    Call gpHabSalvar        'Posiciona controles de salvar e desfazer.
    chkCosseguro(pLinha).Value = True

    cboConSeguradora(pLinha).Enabled = True
    txtConApolice(pLinha).Enabled = True
    txtConApolice(pLinha).Text = "0"
    txtConItem(pLinha).Enabled = True
    txtConItem(pLinha).Text = "0"
    gIndCosseguro = pLinha
    If gCarga = False Then
        If cboConSeguradora(pLinha).Enabled And cboConSeguradora(pLinha).Visible Then
            cboConSeguradora(pLinha).SetFocus
        End If
    End If
End Sub

Private Sub cboConSeguradora_Change(Index As Integer)
    Call gpLimPremios
End Sub
Private Sub chkCosseguro_Click(Index As Integer, Value As Integer)
    Call gpHabSalvar
    Select Case Value
    Case False
        Call mpExcCosseguro(Index)
    Case True
        Call mpIncCosseguro(Index)
    End Select
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

Private Sub Form_Load()
    Dim lbdP0044700 As ADODB.Connection    'Arquivo - P0044700.
    Dim lMensagem As String         'Mensagem.
    Dim lrsP0044700 As clsYasRecordSet  'Registro - P0044700.
    Dim lSelect As String
    '1. Centraliza form.
    Call gM46V999.gpCenForm(Me)

    '2. Carrega combo de seguradoras.
    If gM46V999.gfAbrBasDados("P0044700", gM46V999.gTipP0044700, gM46V999.gEndP0044700, lbdP0044700, lMensagem) = False Then
        Exit Sub
    End If
    lSelect = "SELECT CSeguradora, Seguradora FROM YTAB_SEGURADORA"
    If gM46V999.gfObtRegistro(lbdP0044700, lSelect, lrsP0044700, lMensagem) = False Then
        Call gM46V999.gpFecha2(lbdP0044700)
        Exit Sub
    End If
    While Not lrsP0044700.EOF
        Select Case lrsP0044700!CSeguradora
        Case 641
            'Yasuda - não incluir.
        Case 754, 872, 883
            'Seguradoras sem nome - inválidas - não incluir.
        Case Else
            cboConSeguradora(1).AddItem lrsP0044700!CSeguradora & " - " & lrsP0044700!Seguradora
            cboConSeguradora(1).ItemData(cboConSeguradora(1).NewIndex) = lrsP0044700!CSeguradora
            cboConSeguradora(2).AddItem lrsP0044700!CSeguradora & " - " & lrsP0044700!Seguradora
            cboConSeguradora(2).ItemData(cboConSeguradora(2).NewIndex) = lrsP0044700!CSeguradora
            cboConSeguradora(3).AddItem lrsP0044700!CSeguradora & " - " & lrsP0044700!Seguradora
            cboConSeguradora(3).ItemData(cboConSeguradora(3).NewIndex) = lrsP0044700!CSeguradora
            cboConSeguradora(4).AddItem lrsP0044700!CSeguradora & " - " & lrsP0044700!Seguradora
            cboConSeguradora(4).ItemData(cboConSeguradora(4).NewIndex) = lrsP0044700!CSeguradora
            cboConSeguradora(5).AddItem lrsP0044700!CSeguradora & " - " & lrsP0044700!Seguradora
            cboConSeguradora(5).ItemData(cboConSeguradora(5).NewIndex) = lrsP0044700!CSeguradora
        End Select
        lrsP0044700.MoveNext
    Wend
    Call gM46V999.gpFecha2(lbdP0044700)

End Sub
Private Sub txtConApolice_GotFocus(Index As Integer)
    With txtConApolice(Index)
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtConApolice_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtConApolice_LostFocus(Index As Integer)
    If gM46V999.gfPreenchido(txtConApolice(Index).Text) Then
        txtConApolice(Index).Text = UCase$(txtConApolice(Index).Text)
    End If
End Sub
Private Sub txtConItem_Change(Index As Integer)
    Call gpHabSalvar
End Sub
Private Sub txtConItem_GotFocus(Index As Integer)
    With txtConItem(Index)
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub
Private Sub txtConItem_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

