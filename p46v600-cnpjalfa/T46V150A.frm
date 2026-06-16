VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT46V150A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ajustes"
   ClientHeight    =   2010
   ClientLeft      =   4125
   ClientTop       =   2295
   ClientWidth     =   2895
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   2895
   ShowInTaskbar   =   0   'False
   Begin Threed.SSCommand cmdExtra 
      Height          =   240
      Index           =   1
      Left            =   0
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   240
      _Version        =   65536
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   78
      ForeColor       =   12582912
      Picture         =   "T46V150A.frx":0000
   End
   Begin Threed.SSCommand cmdExtra 
      Height          =   240
      Index           =   2
      Left            =   250
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   240
      _Version        =   65536
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   78
      ForeColor       =   12582912
      Picture         =   "T46V150A.frx":031A
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   735
      Left            =   45
      TabIndex        =   1
      Top             =   945
      Width           =   2760
      _Version        =   65536
      _ExtentX        =   4868
      _ExtentY        =   1296
      _StockProps     =   14
      Caption         =   "Alteração de Prêmio e Prêmio Anual"
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
      Begin Threed.SSCommand cmdAjuste 
         Height          =   240
         Left            =   1890
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   405
         Width           =   240
         _Version        =   65536
         _ExtentX        =   423
         _ExtentY        =   423
         _StockProps     =   78
         ForeColor       =   12582912
         Picture         =   "T46V150A.frx":0634
      End
      Begin Threed.SSOption OptAjustePre 
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   3
         Top             =   405
         Width           =   825
         _Version        =   65536
         _ExtentX        =   1455
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Não"
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
         Value           =   -1  'True
      End
      Begin Threed.SSOption OptAjustePre 
         Height          =   195
         Index           =   1
         Left            =   1125
         TabIndex        =   4
         Top             =   405
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Sim"
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
   Begin Threed.SSFrame SSFrame1 
      Height          =   825
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   2760
      _Version        =   65536
      _ExtentX        =   4868
      _ExtentY        =   1455
      _StockProps     =   14
      Caption         =   "Registra diferença de Prêmios"
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
      Begin Threed.SSCommand cmdRegistra 
         Height          =   240
         Left            =   1890
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   360
         Width           =   240
         _Version        =   65536
         _ExtentX        =   423
         _ExtentY        =   423
         _StockProps     =   78
         ForeColor       =   12582912
         Picture         =   "T46V150A.frx":094E
      End
      Begin Threed.SSOption optDiferenca 
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   2
         Top             =   405
         Width           =   825
         _Version        =   65536
         _ExtentX        =   1455
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Não"
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
         Value           =   -1  'True
      End
      Begin Threed.SSOption optDiferenca 
         Height          =   195
         Index           =   1
         Left            =   1125
         TabIndex        =   5
         Top             =   405
         Width           =   600
         _Version        =   65536
         _ExtentX        =   1058
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Sim"
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
   Begin Threed.SSCommand cmdOk 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   300
      Left            =   990
      TabIndex        =   10
      ToolTipText     =   "Volta à janela de item."
      Top             =   1710
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1764
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Ok"
      ForeColor       =   8388608
   End
End
Attribute VB_Name = "frmT46V150A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    Me.Hide
End Sub

Private Sub cmdRegistra_Click()
    frmT46V149A.Show 1
End Sub

Private Sub Form_Load()
    Call gM46V999.gpCenForm(Me)
    frmT46V100A.panForm.Caption = Mid$(Me.Name, 4)
    cmdAjuste.Enabled = False
    cmdAjuste.Picture = cmdExtra(2).Picture
    cmdRegistra.Enabled = False
    cmdRegistra.Picture = cmdExtra(2).Picture
End Sub


Private Sub OptAjustePre_Click(Index As Integer, Value As Integer)
'1. Posiciona controles.
    Select Case Index
    Case 0  'Não.
        cmdAjuste.Enabled = False
        cmdAjuste.Picture = cmdExtra(2).Picture
    Case 1  'Sim.
        cmdAjuste.Enabled = True
        cmdAjuste.Picture = cmdExtra(1).Picture
    End Select

    '2. Se ação do programa ou processo de carga, encerra processo.
    If mPrograma = True Then
        Exit Sub
    End If
    If gCarga = True Then
        Exit Sub
    End If

End Sub

Private Sub optDiferenca_Click(Index As Integer, Value As Integer)
    Select Case Index
    Case 0  'Não.
        cmdRegistra.Enabled = False
        cmdRegistra.Picture = cmdExtra(2).Picture
    Case 1  'Sim.
        cmdRegistra.Enabled = True
        cmdRegistra.Picture = cmdExtra(1).Picture
    End Select

    'Processa opção do usuário.
    Select Case Index
    Case 0  'Não.
        lMensagem = "AJS0001 - " & gM46V999.gPrefixo & "Excluído registro de Diferença de prêmios. " & _
                    "A Diferença de Prêmio será desconsiderada. Confirma esta opção?"
        If MsgBox(lMensagem, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            Call gM46V999.gpGraLog(0, lMensagem & " Sim.")
            On Error Resume Next
            Unload frmT46V149A
            On Error GoTo 0
            Exit Sub
        End If
        Call gM46V999.gpGraLog(0, lMensagem & " Não.")
        mPrograma = True
        optDiferenca(1).Value = True
        optDiferenca(1).SetFocus
    Case 1  'Sim.
        Call cmdRegistra_Click
    End Select

End Sub


