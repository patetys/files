VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT46V008A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Localização dos dados"
   ClientHeight    =   4620
   ClientLeft      =   465
   ClientTop       =   540
   ClientWidth     =   6105
   ControlBox      =   0   'False
   Icon            =   "T46V008A.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   Begin Threed.SSFrame fraX 
      Height          =   1900
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   2100
      Width           =   6100
      _Version        =   65536
      _ExtentX        =   10760
      _ExtentY        =   3351
      _StockProps     =   14
      Caption         =   " Selecione uma das seguintes opções "
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
      Alignment       =   2
      Begin VB.OptionButton optLocalizacao 
         Caption         =   "&Excluir dados do local de destino e substituir pelos dados do local de origem."
         ForeColor       =   &H00800000&
         Height          =   500
         Index           =   0
         Left            =   100
         TabIndex        =   10
         Top             =   300
         Width           =   5900
      End
      Begin VB.OptionButton optLocalizacao 
         Caption         =   "&Manter dados existentes no local de destino e excluir dados do local de origem."
         ForeColor       =   &H00800000&
         Height          =   500
         Index           =   1
         Left            =   100
         TabIndex        =   9
         Top             =   800
         Value           =   -1  'True
         Width           =   5900
      End
      Begin VB.OptionButton optLocalizacao 
         Caption         =   "&Interromper o processo."
         ForeColor       =   &H00800000&
         Height          =   500
         Index           =   2
         Left            =   100
         TabIndex        =   8
         Top             =   1300
         Width           =   5900
      End
   End
   Begin Threed.SSFrame fraX 
      Height          =   2000
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6100
      _Version        =   65536
      _ExtentX        =   10760
      _ExtentY        =   3528
      _StockProps     =   14
      Caption         =   " Solicitada a seguinte alteração de localização dos dados "
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
      Alignment       =   2
      Begin VB.Label lblX 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Local de origem"
         ForeColor       =   &H00800000&
         Height          =   250
         Index           =   1
         Left            =   100
         TabIndex        =   6
         Top             =   300
         Width           =   5900
      End
      Begin VB.Label lblOrigem 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800000&
         Height          =   250
         Left            =   100
         TabIndex        =   5
         Top             =   550
         Width           =   5900
      End
      Begin VB.Label lblX 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Local de destino"
         ForeColor       =   &H00800000&
         Height          =   250
         Index           =   2
         Left            =   100
         TabIndex        =   4
         Top             =   950
         Width           =   5900
      End
      Begin VB.Label lblDestino 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800000&
         Height          =   250
         Left            =   100
         TabIndex        =   3
         Top             =   1200
         Width           =   5900
      End
      Begin VB.Label lblX 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Já existem dados de cotações e propostas no local de destino."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   250
         Index           =   3
         Left            =   100
         TabIndex        =   2
         Top             =   1600
         Width           =   5900
      End
   End
   Begin Threed.SSCommand cmdOk 
      Height          =   300
      Left            =   2550
      TabIndex        =   0
      Top             =   4200
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1764
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Ok"
      ForeColor       =   8388608
   End
End
Attribute VB_Name = "frmT46V008A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdOK_Click()
    Dim lMensagem As String    'Mensagem.

    If optLocalizacao(0) = True Then
        'Excluir dados do local de destino e substituir pelos dados do local de origem.
        lMensagem = "AGE0044 - Opção selecionada: excluir dados do local de destino e substituir " & _
                    "pelos dados do local de origem." & vbLf & vbLf & _
                    "ESTA OPÇÃO EXCLUIRÁ TODOS OS DADOS EXISTENTES EM " & lblDestino.Caption & "." & _
                    vbLf & vbLf & "ATENÇÃO: num ambiente de rede, o sistema poderá estar excluindo " & _
                    "os dados existentes no servidor." & vbLf & vbLf & "Confirma esta opção?"
        If MsgBox(lMensagem, vbCritical + vbYesNo + vbDefaultButton2) = vbYes Then
            Call gM46V999.gpGraLog(0, lMensagem & " SIM.")
        Else
            Call gM46V999.gpGraLog(0, lMensagem & " NÃO.")
            Exit Sub
        End If
    Else
        If optLocalizacao(1) = True Then
            'Manter dados existentes no local de destino e excluir dados do local de origem.
            lMensagem = "AGE0045 - Opção selecionada: manter dados do local de destino e excluir " & _
                        "dados do local de origem." & vbLf & vbLf & _
                        "ESTA OPÇÃO EXCLUIRÁ TODOS OS DADOS EXISTENTES EM " & lblOrigem.Caption & "." & _
                        vbLf & vbLf & "ATENÇÃO: num ambiente de rede, o sistema poderá estar " & _
                        "excluindo os dados existentes no servidor." & vbLf & vbLf & _
                        "Confirma esta opção?"
            If MsgBox(lMensagem, vbCritical + vbYesNo + vbDefaultButton2) = vbYes Then
                Call gM46V999.gpGraLog(0, lMensagem & " SIM.")
            Else
                Call gM46V999.gpGraLog(0, lMensagem & " NÃO.")
                Exit Sub
            End If
        Else
            If optLocalizacao(2) = True Then
                'Interromper o processo - nada a fazer.
            Else
                Call gM46V999.gpGraLog(1, "FNP0029 - Opção não selecionada.")
            End If
        End If
    End If
    Me.Hide
    Screen.MousePointer = vbHourglass
End Sub
Private Sub Form_Activate()
    frmT46V001A.panForm.Caption = "T28V011A"
End Sub
Private Sub Form_Load()
    Call gM46V999.gpCenForm(Me)
    Screen.MousePointer = vbDefault
End Sub

