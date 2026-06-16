VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT46V146A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selecione opção para transmissão"
   ClientHeight    =   1605
   ClientLeft      =   2100
   ClientTop       =   2700
   ClientWidth     =   6735
   ControlBox      =   0   'False
   Icon            =   "T46V146A.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraOpcao 
      Caption         =   "Opção para transmissão"
      ForeColor       =   &H00C00000&
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   6495
      Begin VB.OptionButton optOpcao 
         Caption         =   "Reemissão"
         ForeColor       =   &H00C00000&
         Height          =   400
         Index           =   4
         Left            =   5160
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optOpcao 
         Caption         =   "Suspender para produto"
         ForeColor       =   &H00C00000&
         Height          =   400
         Index           =   3
         Left            =   2760
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optOpcao 
         Caption         =   "Emitir documento"
         ForeColor       =   &H00C00000&
         Height          =   400
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   1275
      End
      Begin VB.OptionButton optOpcao 
         Caption         =   "Suspender"
         ForeColor       =   &H00C00000&
         Height          =   405
         Index           =   1
         Left            =   1440
         TabIndex        =   1
         Top             =   240
         Width           =   1155
      End
      Begin VB.OptionButton optOpcao 
         Caption         =   "Emitir produto"
         ForeColor       =   &H00C00000&
         Height          =   400
         Index           =   2
         Left            =   4200
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
   End
   Begin Threed.SSCommand cmdOk 
      Height          =   300
      Left            =   2160
      TabIndex        =   3
      Top             =   1200
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1764
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Ok"
      ForeColor       =   12582912
   End
   Begin Threed.SSCommand cmdCancelar 
      Cancel          =   -1  'True
      Height          =   300
      Left            =   3360
      TabIndex        =   4
      Top             =   1200
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1764
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Cancelar"
      ForeColor       =   12582912
   End
End
Attribute VB_Name = "frmT46V146A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Dim lTipoTransmissao    As Long     'Tipo de transmissão
                                    '1 = Emitir
                                    '2 = Suspender
                                    '3 = Automóvel
                                    '5 = Suspender para produto
                                    '6 = Reemissão

Private Sub cmdCancelar_Click()
    lTipoTransmissao = 0
    Unload Me
End Sub
Private Sub cmdOk_Click()
    If optOpcao(0).Value = True Then
        lTipoTransmissao = 1
    ElseIf optOpcao(1).Value = True Then
        frmT46V147A.mpIntanciarForm (gNumPI)
        lTipoTransmissao = 2
    ElseIf optOpcao(2).Value = True Then
        lTipoTransmissao = 3
    ElseIf optOpcao(3).Value = True Then
        lTipoTransmissao = 5
    ElseIf optOpcao(4).Value = True Then
        lTipoTransmissao = 6
    End If
    
    Unload Me
End Sub
Public Sub InstanciarForm(ByRef pTipoTransmissao As Long, Optional pHabilitarCancelar As Boolean = True)
    'Habilita ou não o cancelar
    cmdCancelar.Enabled = pHabilitarCancelar
    
    Me.Show vbModal
    pTipoTransmissao = lTipoTransmissao
End Sub
Private Sub Form_Activate()
    frmT46V100A.panForm.Caption = Mid$(Me.Name, 4)
    If cmdOk.Visible = True And cmdOk.Enabled = True Then
        cmdOk.SetFocus
    End If
End Sub
Private Sub Form_Load()
    'Centralizar form
    Call gpCenForm(Me)

    If gEfetuouCalculo = False Then
        optOpcao(0).Enabled = False
        optOpcao(1).Enabled = True
        optOpcao(2).Enabled = False
        optOpcao(3).Enabled = False
        optOpcao(4).Enabled = False
    Else
        optOpcao(0).Enabled = True
        optOpcao(1).Enabled = True
        optOpcao(2).Enabled = True
        optOpcao(3).Enabled = True
        optOpcao(4).Enabled = True
    End If
    If gPerAcesso <> 1 Then
        optOpcao(2).Enabled = False
    Else
        If gEfetuouCalculo = True Then
            optOpcao(2).Enabled = True
        End If
    End If
End Sub
