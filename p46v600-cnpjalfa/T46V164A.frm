VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT46V164A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Perfil (Homologação)"
   ClientHeight    =   3195
   ClientLeft      =   1710
   ClientTop       =   1050
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   Begin Threed.SSCommand SSCommand2 
      Height          =   330
      Left            =   3825
      TabIndex        =   12
      Top             =   2760
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   582
      _StockProps     =   78
      Caption         =   "Cancelar"
   End
   Begin Threed.SSCommand cmdOk 
      Height          =   330
      Left            =   2655
      TabIndex        =   11
      Top             =   2760
      Width           =   1050
      _Version        =   65536
      _ExtentX        =   1852
      _ExtentY        =   582
      _StockProps     =   78
      Caption         =   "Ok"
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   2445
      Left            =   2520
      TabIndex        =   1
      Top             =   135
      Width           =   2895
      _Version        =   65536
      _ExtentX        =   5106
      _ExtentY        =   4313
      _StockProps     =   14
      Caption         =   "Privilégios"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSCheck ChkTaxa 
         Height          =   240
         Left            =   270
         TabIndex        =   10
         Top             =   1185
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   423
         _StockProps     =   78
         Caption         =   "Alterar Taxa Final"
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
      Begin Threed.SSCheck ChkDesconto 
         Height          =   240
         Left            =   270
         TabIndex        =   9
         Top             =   795
         Width           =   2400
         _Version        =   65536
         _ExtentX        =   4233
         _ExtentY        =   423
         _StockProps     =   78
         Caption         =   "Qualquer desconto comercial"
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
      Begin Threed.SSCheck ChkNorma 
         Height          =   240
         Left            =   270
         TabIndex        =   8
         Top             =   405
         Width           =   1725
         _Version        =   65536
         _ExtentX        =   3043
         _ExtentY        =   423
         _StockProps     =   78
         Caption         =   "Digitar sem Normas"
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
      Begin Threed.SSCheck ChkFranquia 
         Height          =   240
         Left            =   270
         TabIndex        =   14
         Top             =   1575
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   423
         _StockProps     =   78
         Caption         =   "Alterar Franquia"
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
      Begin Threed.SSCheck ChkSiscota 
         Height          =   240
         Left            =   285
         TabIndex        =   15
         Top             =   1965
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   423
         _StockProps     =   78
         Caption         =   "Sem Siscota"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
      End
   End
   Begin Threed.SSFrame Perfil 
      Height          =   2985
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   2310
      _Version        =   65536
      _ExtentX        =   4075
      _ExtentY        =   5265
      _StockProps     =   14
      Caption         =   "Perfil"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSOption optPerfil 
         Height          =   195
         Index           =   6
         Left            =   315
         TabIndex        =   7
         Top             =   2205
         Width           =   1500
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "6 - Operação"
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
      Begin Threed.SSOption optPerfil 
         Height          =   240
         Index           =   5
         Left            =   315
         TabIndex        =   6
         Top             =   1845
         Width           =   1770
         _Version        =   65536
         _ExtentX        =   3122
         _ExtentY        =   423
         _StockProps     =   78
         Caption         =   "5 - Digitação Yasuda"
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
      Begin Threed.SSOption optPerfil 
         Height          =   195
         Index           =   4
         Left            =   315
         TabIndex        =   5
         Top             =   1485
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "4 - Somente Cálculo"
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
      Begin Threed.SSOption optPerfil 
         Height          =   240
         Index           =   3
         Left            =   315
         TabIndex        =   4
         Top             =   1080
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   423
         _StockProps     =   78
         Caption         =   "3 - Gestor de Cálculo"
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
      Begin Threed.SSOption optPerfil 
         Height          =   285
         Index           =   2
         Left            =   315
         TabIndex        =   3
         Top             =   675
         Width           =   1725
         _Version        =   65536
         _ExtentX        =   3043
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "2 - Digitação IST"
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
      Begin Threed.SSOption optPerfil 
         Height          =   285
         Index           =   1
         Left            =   315
         TabIndex        =   2
         Top             =   270
         Width           =   1770
         _Version        =   65536
         _ExtentX        =   3122
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "1 - Gestor Digitação"
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
      Begin Threed.SSOption optPerfil 
         Height          =   195
         Index           =   7
         Left            =   315
         TabIndex        =   13
         Top             =   2520
         Width           =   1500
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "7 - Triagem"
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
End
Attribute VB_Name = "frmT46V164A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Dim objTabUser As clsA46V725A    'Objeto da classe clsA28V713A (Tab_User)
    Dim objstcTabUser As stcA46V725B    'Objeto da classe de estrutura stcA28V713B (Tab_User)
    Dim lMensagem As String
    Dim i As Integer


    Set objTabUser = New clsA46V725A
    Set objTabUser.M46V999 = gM46V999

    Set objstcTabUser = New stcA46V725B
    objstcTabUser.login_user_rede = UCase(gM46V999.gNomUsuario)

    objstcTabUser.prior_m_dcoml = IIf(ChkDesconto.Value = True, 1, 0)
    objstcTabUser.prior_m_normas = IIf(ChkNorma.Value = True, 1, 0)
    objstcTabUser.prior_m_taxa = IIf(ChkTaxa.Value = True, 1, 0)
    objstcTabUser.prior_m_franquia = IIf(ChkFranquia = True, 1, 0)
    ''###PROJETO SISCOTA
    ''DESCOMENTAR LINHAS
    objstcTabUser.prior_m_siscota = IIf(ChkSiscota = True, 1, 0)

    For i = 1 To 7
        If optPerfil(i).Value Then
            objstcTabUser.prior_multi = i
            Exit For
        End If
    Next

    If Not objTabUser.mfAlterar(objstcTabUser) Then
        lMensagem = objTabUser.UltimoErro
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If

    If ChkNorma.Value Then
        gM46V999.VerificarNormas = False
    Else
        gM46V999.VerificarNormas = True
    End If

    Unload Me
End Sub

Private Sub Form_Activate()
    If optPerfil(gM46V999.PerfilAcesso).Enabled Then
        optPerfil(gM46V999.PerfilAcesso).SetFocus
    End If
    optPerfil(gM46V999.PerfilAcesso).Value = True
End Sub

Private Sub Form_Load()
    Dim objTabUser As clsA46V725A    'Objeto da classe clsA28V713A (Tab_User)
    Dim objstcTabUser As stcA46V725B    'Objeto da classe de estrutura stcA28V713B (Tab_User)
    Dim lMensagem As String
    Dim lI As Integer

    Call gM46V999.gpCenForm(Me)


    Set objTabUser = New clsA46V725A
    Set objTabUser.M46V999 = gM46V999

    Set objstcTabUser = New stcA46V725B
    objstcTabUser.login_user_rede = UCase(gM46V999.gNomUsuario)
    If Not objTabUser.mfCarregarColecao(objstcTabUser) Then
        lMensagem = objTabUser.UltimoErro
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If

    If objTabUser.Colecao.Count = 0 Then
        cmdOk.Enabled = False
        Exit Sub
    End If

    Set objstcTabUser = objTabUser.Colecao.Item(1)

    ChkDesconto.Value = IIf(objstcTabUser.prior_m_dcoml = 1, True, False)
    ChkNorma.Value = IIf(objstcTabUser.prior_m_normas = 1, True, False)
    ChkTaxa.Value = IIf(objstcTabUser.prior_m_taxa = 1, True, False)
    ChkFranquia.Value = IIf(objstcTabUser.prior_m_franquia = 1, True, False)
    '    '###PROJETO SISCOTA
    '    'DESCOMENTAR LINHAS
    ChkSiscota.Value = IIf(objstcTabUser.prior_m_siscota = 1, True, False)

    If objstcTabUser.cod_Depto = 944 And UCase(gM46V999.gNomUsuario) <> "SHIBA" Then
        For lI = 1 To 7
            optPerfil(lI).Enabled = False
        Next
        ChkDesconto.Enabled = False
        ChkNorma.Enabled = False
        ChkTaxa.Enabled = False
        ChkFranquia.Enabled = False
        ChkSiscota.Enabled = False
    End If

End Sub


Private Sub SSCommand1_Click()

End Sub

Private Sub optPerfil_Click(Index As Integer, Value As Integer)
    Select Case Index
    Case 1, 2, 5, 7
        'ChkTaxa.Value = False
        'ChkTaxa.Enabled = False
    Case Else
        ChkTaxa.Enabled = True
    End Select
End Sub

Private Sub SSCommand2_Click()
    Unload Me
End Sub


