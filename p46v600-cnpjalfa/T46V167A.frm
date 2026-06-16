VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT46V167A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mover Desconto Comercial"
   ClientHeight    =   1155
   ClientLeft      =   5445
   ClientTop       =   2880
   ClientWidth     =   2790
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1155
   ScaleWidth      =   2790
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtDescComercial 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   495
      MaxLength       =   5
      TabIndex        =   0
      ToolTipText     =   "Número de bloqueto do pagamento antecipado."
      Top             =   345
      Width           =   1770
   End
   Begin Threed.SSCommand cmdOk 
      Height          =   300
      Left            =   360
      TabIndex        =   1
      ToolTipText     =   "Volta à proposta."
      Top             =   765
      Width           =   960
      _Version        =   65536
      _ExtentX        =   1693
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Mover"
      ForeColor       =   8388608
   End
   Begin Threed.SSCommand cmdCancel 
      Cancel          =   -1  'True
      Height          =   300
      Left            =   1485
      TabIndex        =   3
      ToolTipText     =   "Volta à proposta."
      Top             =   765
      Width           =   960
      _Version        =   65536
      _ExtentX        =   1693
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Cancelar"
      ForeColor       =   8388608
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Desconto Comercial"
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   26
      Left            =   495
      TabIndex        =   2
      ToolTipText     =   "Número de bloqueto do pagamento antecipado."
      Top             =   90
      Width           =   1770
   End
End
Attribute VB_Name = "frmT46V167A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    cmdOk.Tag = ""
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    Dim lobjstcPedLoc As stcA46V708B    'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
    Dim lobjstcPedCobert As stcA46V704B
    Dim lObjPedCobert As clsA46V704A    'Objeto da classe de coberturas.
    Dim lobjCtrlEmis As clsA46V712A    'Objeto da classe clsA46V712A (TAB_CTRL_EMIS)
    Dim lChave As Long    'Posição na coleção.
    Dim lColecaoTemp As Collection

    Set lObjPedCobert = New clsA46V704A
    Set lObjPedCobert.M46V999 = gM46V999


    Set lobjCtrlEmis = New clsA46V712A
    Set lobjCtrlEmis.M46V999 = gM46V999

    Set lobjstcPedCobert = New stcA46V704B
    Set lobjstcPedCobert.M46V999 = gM46V999

    Set lobjstcPedLoc = New stcA46V708B
    Set lobjstcPedLoc.M46V999 = gM46V999

    With frmT46V101A
        If .objstcPedido.ColecaoPedLoc.Count > 0 Then
            Set lColecaoTemp = New Collection
            For Each lobjstcPedLoc In .objstcPedido.ColecaoPedLoc
                For Each lobjstcPedCobert In lobjstcPedLoc.ColecaoPedCobert
                    lobjstcPedCobert.Per_Desc_Coml = Me.txtDescComercial
                    If Not lObjPedCobert.mfAlterarDescontoComercial(lobjstcPedCobert, "", "") Then
                        Call gM46V999.gpGraLog(1, "FGE00167 - Mover Desconto " & lObjPedCobert.UltimoErro)
                        Exit Sub
                    End If
                    If lobjstcPedCobert.Num_Item = gM46V999.CDblx(.txtNumItem.Text) Then
                        lColecaoTemp.Add lobjstcPedCobert, lobjstcPedCobert.Class_ID
                    End If
                Next
            Next
            Set lobjstcPedLoc = Nothing
            Set lobjstcPedLoc = .objstcPedido.ColecaoPedLoc.Item(.objstcPedido.NOSSO_NUMERO & Format(.txtNumItem.Text, "00000"))
            Set lobjstcPedLoc.ColecaoPedCobert = Nothing
            Set lobjstcPedLoc.ColecaoPedCobert = lColecaoTemp

            .objstcPedido.ColecaoPedLoc.Remove lobjstcPedLoc.Class_ID
            'Verifico em que lugar da coleção será adicionado.
            'É feito isso para que não seja exibido com os número fora de ordem
            lChave = gpPosicaoColecao(lobjstcPedLoc.Num_Item)
            If .objstcPedido.ColecaoPedLoc.Count > 0 And .objstcPedido.ColecaoPedLoc.Count <> lChave - 1 Then
                .objstcPedido.ColecaoPedLoc.Add lobjstcPedLoc, lobjstcPedLoc.Class_ID, lChave
            Else
                .objstcPedido.ColecaoPedLoc.Add lobjstcPedLoc, lobjstcPedLoc.Class_ID
            End If

            .objstcPedido.objstcCtrlEmis.Per_Desc_Coml = Me.txtDescComercial
            If Not lobjCtrlEmis.mfAlterarDescontoComercial(.objstcPedido.objstcCtrlEmis, "", "") Then
                Call gM46V999.gpGraLog(1, "FGE00168 - Mover Desconto " & lobjCtrlEmis.UltimoErro)
                Exit Sub
            End If


        End If
        cmdOk.Tag = "ALTERADO"
        Me.Hide
    End With
End Sub
Private Sub Form_Activate()
    frmT46V100A.panForm.Caption = Mid$(Me.Name, 4)
    cmdOk.Tag = ""
Finalizar:
    If gM46V999.gDesabilitarForms Or gM46V999.gDesabFormsCotacao Then
        If Me.Tag <> "Desabilitado" Then
            DesabilitaControles Me
            cmdCancel.Enabled = True
        End If
    Else
        If Me.Tag = "Desabilitado" Then
            HabilitaControles Me
        End If
    End If

End Sub
Private Sub Form_Load()
    Call gM46V999.gpCenForm(Me)
    txtDescComercial.Text = frmT46V101A.txtDescComercial.Text
End Sub

Private Sub txtDescComercial_GotFocus()
    With txtDescComercial
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub
Private Sub txtDescComercial_KeyPress(KeyAscii As Integer)
    KeyAscii = ConsisteKeyNum(KeyAscii, ",")
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
    Call gpHabSalvar
End Sub
Private Sub txtDescComercial_LostFocus()
    With txtDescComercial
        If gM46V999.gfPreenchido(.Text) Then
            If IsNumeric(.Text) Then
                .Text = Format$(CDbl(.Text), "standard")
            Else
                .Text = Format$("0,00", "standard")
            End If
        Else
            .Text = Format$("0,00", "standard")
        End If
    End With
End Sub

Public Sub mpInstanciarForm(pValor As Double)
    txtDescComercial.Text = pValor
    Call cmdOK_Click
End Sub


