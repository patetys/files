VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT47V020A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selecione opção para transmissão"
   ClientHeight    =   1425
   ClientLeft      =   4770
   ClientTop       =   2820
   ClientWidth     =   4530
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraOpcao 
      Caption         =   "Opção para transmissão"
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4335
      Begin VB.OptionButton optOpcao 
         Caption         =   "Suspender para produto"
         ForeColor       =   &H00800000&
         Height          =   400
         Index           =   2
         Left            =   2760
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optOpcao 
         Caption         =   "Emitir documento"
         ForeColor       =   &H00800000&
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
         ForeColor       =   &H00800000&
         Height          =   405
         Index           =   1
         Left            =   1440
         TabIndex        =   1
         Top             =   240
         Width           =   1155
      End
   End
   Begin Threed.SSCommand cmdOk 
      Height          =   300
      Left            =   1050
      TabIndex        =   2
      Top             =   1110
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
      Left            =   2235
      TabIndex        =   3
      Top             =   1110
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1764
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Cancelar"
      ForeColor       =   8388608
   End
End
Attribute VB_Name = "frmT47V020A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lTipoTransmissao As Long        'Tipo de transmissão
                                    '1 = Emitir
                                    '2 = Suspender
                                    '3 = Automóvel
                                    '5 = Suspender para produto
                                    '6 = Reemissão
Dim lStatusAcao    As Integer

Private Sub cmdCancelar_Click()
    lTipoTransmissao = 0
    Unload Me
End Sub

Private Sub cmdOK_Click()

Dim lMensagem       As String

    lStatusAcao = 0
    
    If optOpcao(0).Value = True Then
        lTipoTransmissao = 1
    ElseIf optOpcao(1).Value = True Then
       ' frmT28V147A.mpIntanciarForm (gNumPI)
                   
        If Not mfVerificarTabCtrlEmisGed(lStatusAcao) Then
            Call gM46V999.gpGraLog(1, lMensagem)
            Exit Sub
        End If
       
        If lStatusAcao = 0 Then
            frmT47V027A.mpInstanciarForm lStatusAcao
        End If

        If lStatusAcao = 0 Then
            lTipoTransmissao = 0
        Else
            lTipoTransmissao = 2
        End If
    ElseIf optOpcao(2).Value = True Then
        lTipoTransmissao = 3
        
    ElseIf optOpcao(3).Value = True Then
'        frmT28V147A.mpIntanciarForm (gNumPI)
        lTipoTransmissao = 5
    ElseIf optOpcao(4).Value = True Then
        lTipoTransmissao = 6
    End If
    
    Unload Me
End Sub

Public Sub mpInstanciarForm(ByRef pTipoTransmissao As Long, ByRef pAcao As Integer, Optional pHabilitarCancelar As Boolean = True)
    'Habilita ou não o cancelar
    cmdCancelar.Enabled = pHabilitarCancelar
    
    Me.Show vbModal
    pTipoTransmissao = lTipoTransmissao
    pAcao = lStatusAcao
End Sub

Private Sub Form_Activate()
    If cmdOk.Visible = True And cmdOk.Enabled = True Then
        cmdOk.SetFocus
    End If
End Sub

Private Sub Form_Load()
    'Centralizar form
    Call gM46V999.gpCenForm(Me)

'Ficha 900591
    If gM46V999.gPodeEmitir = False Then gM46V999.gEfetuouCalculo = False
    
    If gM46V999.gEfetuouCalculo = False Then
        optOpcao(0).Enabled = False
        optOpcao(1).Enabled = True
        optOpcao(2).Enabled = True
        optOpcao(2).Value = True
    Else
        optOpcao(0).Enabled = True
        optOpcao(1).Enabled = True
        optOpcao(2).Enabled = True
        optOpcao(0).Value = True
    End If
'    If gM46V999.PerfilAcesso <> e_Perfil_DIG_Gestor Then
'        optOpcao(2).Enabled = False
'    Else
'        If gM46V999.gEfetuouCalculo = True Then
'            optOpcao(2).Enabled = True
'        End If
'    End If
End Sub
Private Function mfVerificarTabCtrlEmisGed(pStatusAcao As Integer) As Boolean

Dim lSql            As String
Dim lStatusAcao     As Integer
Dim lrst            As clsYasRecordSet
Dim lMensagem       As String
Dim bdGed           As ADODB.Connection

    mfVerificarTabCtrlEmisGed = False
    
    If Not gM46V999.gfAbrBasDados("GED", gM46V999.gTipGed, gM46V999.gEndGed, bdGed, lMensagem) Then
        Call gM46V999.gpGraLog(1, lMensagem)
        Exit Function
    End If
    
    lSql = "Select * from ged.dbo.TAB_CTRL_EMIS "
    lSql = lSql & " WHERE num_pi = " & gM46V111.gNum_Pi
    
    If Not gM46V999.gfObtRegistro(bdGed, lSql, lrst, lMensagem) Then
        Call gM46V999.gpGraLog(1, lMensagem)
        Exit Function
    End If
        
    If lrst.EOF Then 'se nao existir tab_ctrl_emis não chama a tela de motivos
        pStatusAcao = 2
    End If
    
    Call gM46V999.gpFecha2(bdGed)
    mfVerificarTabCtrlEmisGed = True
    Exit Function

End Function


