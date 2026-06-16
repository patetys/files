VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT46V109A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carnê"
   ClientHeight    =   1005
   ClientLeft      =   2700
   ClientTop       =   5310
   ClientWidth     =   5130
   ControlBox      =   0   'False
   Icon            =   "T46V109A.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1005
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboCarBanco 
      ForeColor       =   &H00800000&
      Height          =   315
      ItemData        =   "T46V109A.frx":000C
      Left            =   2800
      List            =   "T46V109A.frx":000E
      Sorted          =   -1  'True
      TabIndex        =   0
      Text            =   "cboCarBanco"
      ToolTipText     =   "Banco para emissão de carnê."
      Top             =   100
      Width           =   2200
   End
   Begin Threed.SSCommand cmdOk 
      Cancel          =   -1  'True
      Height          =   300
      Left            =   2055
      TabIndex        =   1
      ToolTipText     =   "Volta à proposta."
      Top             =   600
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Ok"
      ForeColor       =   8388608
   End
   Begin VB.Label lblX 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Banco: "
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   0
      Left            =   100
      TabIndex        =   2
      Top             =   100
      Width           =   2700
   End
End
Attribute VB_Name = "frmT46V109A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cboCarBanco_Click()
    Call gpHabSalvar
End Sub
Private Sub cboCarBanco_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        'KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub cboCarBanco_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyBack Then     'Teclou backspace.
        cboCarBanco.ListIndex = -1
    End If
End Sub
Private Sub cmdOk_Click()
    If cboCarBanco.ListIndex = -1 Then
        MsgBox "Favor informar um Banco para gerar Carnê.", vbInformation
        cboCarBanco.SetFocus
    Else
        Me.Hide
    End If
End Sub
Private Sub Form_Activate()
    frmT46V100A.panForm.Caption = Mid$(Me.Name, 4)
    If cboCarBanco.Visible = True And cboCarBanco.Enabled = True Then
        cboCarBanco.SetFocus
        'cboCarBanco.Enabled = False 'Boleto Santsnder
    End If
End Sub
Private Sub Form_Load()
    Dim lMensagem As String    'Mensagem.
    Dim i As Integer
    Dim lDatEmiss As Long
'1. Centraliza form.
    Call gM46V999.gpCenForm(Me)

    '2. Carrega combo de bancos.
    If gfCarCboBancos(2, cboCarBanco, lMensagem, frmT46V101A.objstcPedido.objstcCtrlEmis.Cod_Corr) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If

    Select Case gM46V999.PerfilAcesso
    Case 1, 2, 5
        cboCarBanco.AddItem "215"
        cboCarBanco.ItemData(cboCarBanco.NewIndex) = 215
    End Select
'Boleto Santander
ConsultaSegregacao:
    For i = 0 To cboCarBanco.ListCount - 1
        '2016-08-22: Alterar Itau p Santander
        If frmT46V101A.objstcPedido.objstcCtrlEmis.Cod_Corr = 15701 Then 'mercedes
            If cboCarBanco.ItemData(i) = 33 Then
                cboCarBanco.ListIndex = i
                Exit For
            End If
        Else
            If cboCarBanco.ItemData(i) = frmT46V101A.objstcPedido.Cod_Bco_Cobr Then
                cboCarBanco.ListIndex = i
                cboCarBanco.Enabled = False
                Exit For
            End If
        End If
        '2016-08-22
    Next i
    
    lDatEmiss = 0
    If cboCarBanco.ListIndex = -1 Then
        With frmT46V101A
            
            If .objstcPedLoc.Tip_Emissao <= 200 Then
                lDatEmiss = 0
            Else
                gM46V999.gNum_Apolice = .objstcPedido.Num_Apol_An
                lDatEmiss = gM46V999.gDataEmis_Apol
            End If
        
            If Left(.objstcPedido.objstcCtrlEmis.Num_Bloq, 1) <> 5 Then
                gM46V999.gCod_Ramo = .objstcPedido.Cod_Ramo
                .objstcPedido.Cod_Bco_Cobr = gM46V999.gObtemBoletoSegregacao(.objstcPedLoc.Tip_Emissao, _
                                                                            .objstcPedido.Num_Cgc_Cpf, _
                                                                            .objstcPedido.Cod_Produto, _
                                                                            .objstcPedido.Cod_Segurado, _
                                                                            .objstcPedido.Num_Apol_An, _
                                                                            .objstcPedido.objstcCtrlEmis.Cod_Corr, _
                                                                            lDatEmiss, .objstcPedido.Cod_Tip_Doc)
            End If
            If .objstcPedido.Cod_Bco_Cobr <> 0 Then
                GoTo ConsultaSegregacao
            End If
        End With
    End If
    
End Sub
