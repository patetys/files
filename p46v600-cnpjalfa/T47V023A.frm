VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT47V023A 
   Caption         =   "Cláusulas"
   ClientHeight    =   3750
   ClientLeft      =   2625
   ClientTop       =   2400
   ClientWidth     =   6930
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3750
   ScaleWidth      =   6930
   Begin Threed.SSCommand cmdOk 
      Default         =   -1  'True
      Height          =   300
      Left            =   2963
      TabIndex        =   0
      ToolTipText     =   "Volta à proposta."
      Top             =   3420
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Ok"
      ForeColor       =   8388608
   End
   Begin Threed.SSPanel panRenCongenere 
      Height          =   3225
      Left            =   15
      TabIndex        =   1
      Top             =   90
      Width           =   6900
      _Version        =   65536
      _ExtentX        =   12171
      _ExtentY        =   5689
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
      Begin VB.ComboBox cboClausulas 
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   9
         Left            =   90
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         ToolTipText     =   "Cláusulas"
         Top             =   2790
         Width           =   6735
      End
      Begin VB.ComboBox cboClausulas 
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   8
         Left            =   90
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         ToolTipText     =   "Cláusulas"
         Top             =   2445
         Width           =   6735
      End
      Begin VB.ComboBox cboClausulas 
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   7
         Left            =   90
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         ToolTipText     =   "Cláusulas"
         Top             =   2115
         Width           =   6735
      End
      Begin VB.ComboBox cboClausulas 
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   1
         Left            =   90
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         ToolTipText     =   "Cláusulas"
         Top             =   90
         Width           =   6735
      End
      Begin VB.ComboBox cboClausulas 
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   2
         Left            =   90
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Cláusulas"
         Top             =   420
         Width           =   6735
      End
      Begin VB.ComboBox cboClausulas 
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   3
         Left            =   90
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Cláusulas"
         Top             =   765
         Width           =   6735
      End
      Begin VB.ComboBox cboClausulas 
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   4
         Left            =   90
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Cláusulas"
         Top             =   1095
         Width           =   6735
      End
      Begin VB.ComboBox cboClausulas 
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   5
         Left            =   90
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Cláusulas"
         Top             =   1440
         Width           =   6735
      End
      Begin VB.ComboBox cboClausulas 
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   6
         Left            =   90
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Cláusulas"
         Top             =   1770
         Width           =   6735
      End
   End
End
Attribute VB_Name = "frmT47V023A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cboClausulas_Change(Index As Integer)
    If gCarga Then Exit Sub
    Call gpLimPremios
End Sub

Private Sub cboClausulas_Click(Index As Integer)
    cboClausulas_Change Index
End Sub


Private Sub cmdOK_Click()
Dim i                     As Integer
Dim j                     As Integer

    'verifico se há duplicidade
    For i = 1 To 9
        If cboClausulas(i).ListIndex > -1 Then
            If cboClausulas(i).ItemData(cboClausulas(i).ListIndex) <> 0 Then
                For j = i To 9
                    If i <> j Then
                        If cboClausulas(j).ListIndex > -1 Then
                            If cboClausulas(i).ItemData(cboClausulas(i).ListIndex) = _
                               cboClausulas(j).ItemData(cboClausulas(j).ListIndex) Then
                                 Call gM46V999.gpGraLog(1, "FGE0092 - Cláusula inválida. Duplicidade de Cláusulas.")
                                 Exit Sub
                             End If
                        End If
                    End If
                Next j
            End If
        End If
    Next i
    
    Me.Hide
End Sub
Private Sub Form_Activate()
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
Dim lbdTab_Clau   As ADODB.Connection     'Arquivo: P00Tarifa.
Dim lMensagem     As String               'Mensagem.
Dim lrsTab_Clau   As clsYasRecordSet      'Registro : Tab_Clau.
Dim i             As Integer
Dim lSelect         As String

    '1. Posiciona form.
    Call gM46V999.gpCenForm(Me)
    
    If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTab_Clau, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If
    

    lSelect = " Select *  from Tab_Clau "
    lSelect = lSelect & " Where cod_Ramo = " & gM46V999.gCod_Ramo
    If gM46V999.gCod_Ramo = e_CodRamo_Empresarial Then
        lSelect = lSelect & " and Cod_Clau > 2 "
    End If
    lSelect = lSelect & " order by cod_clau"
    
    If gM46V999.gfObtRegistro(lbdTab_Clau, lSelect, lrsTab_Clau, lMensagem) = False Then
        Call gM46V999.gpFecha2(lbdTab_Clau)
        Call gM46V999.gpGraLog(2, lMensagem)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    For i = 1 To 9
        cboClausulas(i).AddItem " "
        cboClausulas(i).ItemData(cboClausulas(i).NewIndex) = 0
    Next
    Do While Not lrsTab_Clau.EOF
        For i = 1 To 9
            cboClausulas(i).AddItem Format(lrsTab_Clau("COD_CLAU"), "000") & " - " & Trim(lrsTab_Clau("DSC_CLAU"))
            cboClausulas(i).ItemData(cboClausulas(i).NewIndex) = lrsTab_Clau("COD_CLAU")
        Next
        lrsTab_Clau.MoveNext
    Loop
    
    Call gM46V999.gpFecha2(lbdTab_Clau)

End Sub

