VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmT46V200A 
   Caption         =   "Editar Coberturas"
   ClientHeight    =   4680
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4680
   ScaleWidth      =   9810
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtVal_IS 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1590
      MaxLength       =   14
      TabIndex        =   6
      Text            =   "500.000.000,00"
      Top             =   1470
      Width           =   1335
   End
   Begin VB.TextBox txtNum_PI_Meses 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4260
      MaxLength       =   2
      TabIndex        =   8
      Text            =   "12"
      Top             =   1470
      Width           =   375
   End
   Begin VB.Frame fraFranquia 
      Caption         =   "Franquia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2715
      Left            =   60
      TabIndex        =   9
      Top             =   1890
      Width           =   9645
      Begin VB.ComboBox cboDesc_Franq 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "T46V200A.frx":0000
         Left            =   5880
         List            =   "T46V200A.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   300
         Width           =   705
      End
      Begin VB.ComboBox cboMultiploFranquia 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "T46V200A.frx":0015
         Left            =   8100
         List            =   "T46V200A.frx":001F
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   300
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.TextBox txtObs_Franquia 
         Appearance      =   0  'Flat
         Height          =   1455
         Left            =   1710
         MaxLength       =   8000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Text            =   "T46V200A.frx":002B
         Top             =   1080
         Width           =   7815
      End
      Begin VB.TextBox txtTexto_Franquia 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1710
         TabIndex        =   19
         Top             =   690
         Width           =   7815
      End
      Begin VB.TextBox txtVal_Franq_min 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3330
         TabIndex        =   13
         Text            =   "50.000,00"
         Top             =   300
         Width           =   1185
      End
      Begin VB.TextBox txtPer_Franquia 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1710
         MaxLength       =   2
         TabIndex        =   11
         Text            =   "99"
         Top             =   300
         Width           =   345
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Desconto : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4830
         TabIndex        =   14
         Top             =   300
         Width           =   1035
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Multiplo : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   7110
         TabIndex        =   16
         Top             =   300
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Observação : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   180
         TabIndex        =   20
         Top             =   1080
         Width           =   1515
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Texto : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   180
         TabIndex        =   18
         Top             =   690
         Width           =   1515
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Val Min : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2340
         TabIndex        =   12
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Percentual % : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   180
         TabIndex        =   10
         Top             =   300
         Width           =   1515
      End
   End
   Begin VB.TextBox txtDescricao 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   1590
      TabIndex        =   4
      Text            =   "1102"
      Top             =   1110
      Width           =   8025
   End
   Begin VB.TextBox txtCodigo 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   1590
      TabIndex        =   2
      Text            =   "1102"
      Top             =   750
      Width           =   555
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9810
      _ExtentX        =   17304
      _ExtentY        =   1164
      ButtonWidth     =   1349
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgLista"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salvar"
            Key             =   "Salvar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Anterior"
            Key             =   "Anterior"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Próximo"
            Key             =   "Proximo"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair"
            Key             =   "Sair"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgLista 
      Left            =   7290
      Top             =   450
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V200A.frx":1190
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V200A.frx":14E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V200A.frx":1834
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V200A.frx":1B86
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Val IS : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   60
      TabIndex        =   5
      Top             =   1470
      Width           =   1515
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PI Meses : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   3120
      TabIndex        =   7
      Top             =   1470
      Width           =   1125
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Descrição : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   60
      TabIndex        =   3
      Top             =   1110
      Width           =   1515
   End
   Begin VB.Label lblCodigo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Código : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   750
      Width           =   1515
   End
End
Attribute VB_Name = "frmT46V200A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private colCobert As Collection
Private intPosicao As Integer

Public Sub CarregarForm(ByRef pcolCobert As Collection, ByVal pPosicao As Integer)
    Set colCobert = pcolCobert
    intPosicao = pPosicao
    Call Posiciona
    Me.Show vbModal
End Sub

Private Function Salvar() As Boolean
    On Error GoTo Erro

    Salvar = False
    txtTexto_Franquia.Text = Empty
    txtTexto_Franquia.Text = gM46V999.gfTextoFranquia(gM46V999.CDblx(txtVal_Franq_min.Text), gM46V999.CDblx(txtPer_Franquia.Text), IIf(Len(txtObs_Franquia.Text) > 0, "*", ""), txtCodigo.Text, gM46V999.CDblx(txtNum_PI_Meses.Text))

    With colCobert(intPosicao)
        .Val_IS = gM46V999.CDblx(txtVal_IS.Text)
        .Per_Franq = txtPer_Franquia.Text
        .Val_Franq_Min = gM46V999.CDblx(txtVal_Franq_min.Text)
        If .ColecaoNovaFranquia.Count > 0 Then
            .ColecaoNovaFranquia(1).Multiplo_Franquia = IIf(cboMultiploFranquia.Text = "1x", "1", "3")
            .ColecaoNovaFranquia(1).Texto_Observacoes = txtObs_Franquia.Text
        End If
        .Per_Desc_Franq = cboDesc_Franq.Text
        .Num_Meses_PI = txtNum_PI_Meses.Text
    End With
    
    Salvar = True
    Exit Function
Erro:
    MsgBox Err.Number & " - " & Err.Description
End Function

Private Function Posiciona() As Boolean
    On Error GoTo Erro
    
    Posiciona = False
    With colCobert(intPosicao)
        txtCodigo.Text = .Cod_Cobert
        txtDescricao.Text = .gNomeCobertura
        txtVal_IS.Text = Format(.Val_IS, "###,###,##0.00")
        txtPer_Franquia.Text = .Per_Franq
        txtVal_Franq_min.Text = Format(.Val_Franq_Min, "###,###,##0.00")
        If .ColecaoNovaFranquia.Count = 0 Then
           cboMultiploFranquia.Text = "1x"
           txtObs_Franquia.Text = Empty
        Else
            cboMultiploFranquia.Text = .ColecaoNovaFranquia(1).Multiplo_Franquia & "x"
            txtObs_Franquia.Text = .ColecaoNovaFranquia(1).Texto_Observacoes
        End If
        cboDesc_Franq.Text = IIf(.Per_Desc_Franq > 0, "25", "0")
        
        If mfBuscaPerIndenitario(.Cod_Cobert) Then
            txtNum_PI_Meses.Enabled = True
            txtNum_PI_Meses.Text = .Num_Meses_PI
        Else
            txtNum_PI_Meses.Enabled = False
            txtNum_PI_Meses.Text = 0
        End If
    End With
    
    txtTexto_Franquia.Text = Empty
    txtTexto_Franquia.Text = gM46V999.gfTextoFranquia(gM46V999.CDblx(txtVal_Franq_min.Text), gM46V999.CDblx(txtPer_Franquia.Text), IIf(Len(txtObs_Franquia.Text) > 0, "*", ""), txtCodigo.Text, gM46V999.CDblx(txtNum_PI_Meses.Text))
    
    Posiciona = True
    Exit Function
Erro:
    MsgBox Err.Number & " - " & Err.Description
End Function

Private Sub Anterior()

    If Not Salvar Then Exit Sub
    
    intPosicao = intPosicao - 1
    If intPosicao < 1 Then
        intPosicao = 1
    Else
    Call Posiciona
    End If
    
End Sub

Private Sub Proximo()

    If Not Salvar Then Exit Sub
    
    intPosicao = intPosicao + 1
    If intPosicao > colCobert.Count Then
        intPosicao = colCobert.Count
    Else
    Call Posiciona
    End If
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Salvar"
            If Salvar Then
                Unload Me
            End If
        Case "Anterior"
            Call Anterior
        Case "Proximo"
            Call Proximo
        Case "Sair"
            Unload Me
    End Select
End Sub

Private Sub txtObs_Franquia_LostFocus()
    txtTexto_Franquia.Text = Empty
    txtTexto_Franquia.Text = gM46V999.gfTextoFranquia(gM46V999.CDblx(txtVal_Franq_min.Text), gM46V999.CDblx(txtPer_Franquia.Text), IIf(Len(txtObs_Franquia.Text) > 0, "*", ""), txtCodigo.Text, gM46V999.CDblx(txtNum_PI_Meses.Text))
End Sub

Private Sub txtVal_IS_GotFocus()
    With txtVal_IS
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtVal_IS_KeyPress(KeyAscii As Integer)
    KeyAscii = ConsisteKeyNum(KeyAscii, ",")

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtVal_IS_LostFocus()
    Dim i As Integer
    Dim lI As Integer

    With txtVal_IS
        If gM46V999.gfPreenchido(.Text) Then
            If IsNumeric(.Text) Then
                .Text = Format(.Text, "standard")
                If .Text > 2147483647 Then
                    MsgBox "Valor em risco excedeu o limite!", vbInformation, "Atenção"
                    txtVal_IS.SetFocus
                End If
            Else
                .Text = Format(0, "standard")
            End If
        Else
            .Text = Format(0, "standard")
        End If
    End With
End Sub

Private Sub txtVal_Franq_min_GotFocus()
    With txtVal_Franq_min
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtVal_Franq_min_KeyPress(KeyAscii As Integer)
    KeyAscii = ConsisteKeyNum(KeyAscii, ",")

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
    
End Sub

Private Sub txtVal_Franq_min_LostFocus()
    Dim i As Integer
    Dim lI As Integer

    With txtVal_Franq_min
        If gM46V999.gfPreenchido(.Text) Then
            If IsNumeric(.Text) Then
                .Text = Format(.Text, "standard")
                If .Text > 2147483647 Then
                    MsgBox "Valor em risco excedeu o limite!", vbInformation, "Atenção"
                    txtVal_Franq_min.SetFocus
                End If
            Else
                .Text = Format(0, "standard")
            End If
        Else
            .Text = Format(0, "standard")
        End If
    End With
    
    txtTexto_Franquia.Text = Empty
    txtTexto_Franquia.Text = gM46V999.gfTextoFranquia(gM46V999.CDblx(txtVal_Franq_min.Text), gM46V999.CDblx(txtPer_Franquia.Text), IIf(Len(txtObs_Franquia.Text) > 0, "*", ""), txtCodigo.Text, gM46V999.CDblx(txtNum_PI_Meses.Text))
End Sub

Private Sub txtNum_PI_Meses_KeyPress(KeyAscii As Integer)
    KeyAscii = ConsisteKeyNum(KeyAscii, ",")

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtNum_PI_Meses_LostFocus()
    txtTexto_Franquia.Text = Empty
    txtTexto_Franquia.Text = gM46V999.gfTextoFranquia(gM46V999.CDblx(txtVal_Franq_min.Text), gM46V999.CDblx(txtPer_Franquia.Text), IIf(Len(txtObs_Franquia.Text) > 0, "*", ""), txtCodigo.Text, gM46V999.CDblx(txtNum_PI_Meses.Text))
End Sub

Private Sub txtPer_Franquia_KeyPress(KeyAscii As Integer)
    KeyAscii = ConsisteKeyNum(KeyAscii, ",")

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtPer_Franquia_LostFocus()
    txtTexto_Franquia.Text = Empty
    txtTexto_Franquia.Text = gM46V999.gfTextoFranquia(gM46V999.CDblx(txtVal_Franq_min.Text), gM46V999.CDblx(txtPer_Franquia.Text), IIf(Len(txtObs_Franquia.Text) > 0, "*", ""), txtCodigo.Text, gM46V999.CDblx(txtNum_PI_Meses.Text))
End Sub

Private Function mfBuscaPerIndenitario(ByVal pCodCobert As Integer) As Boolean
    Dim lbdTbm_Coef_PI As ADODB.Connection      'Arquivo: P00Multi.
    Dim lMensagem As String                     'Mensagem.
    Dim lrsTbm_Coef_PI As clsYasRecordSet       'Registro : Tbm_Coef_PI.
    Dim lSelect As String
    mfBuscaPerIndenitario = False

    If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTbm_Coef_PI, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Function
    End If

    '1 - Procura pela Cobertura na Tabela de Coeficiente
    'Se encontrar é permitido Cadastrar

    lSelect = " Select *  from Tbm_Coef_PI "
    lSelect = lSelect & " Where Cod_Ramo = " & gM46V999.gCod_Ramo
    lSelect = lSelect & " and Cod_Cobert = " & pCodCobert
    lSelect = lSelect & " and Cod_Tarif = " & gM46V999.gCodProduto
    lSelect = lSelect & " and Cod_Sequencia <= " & gM46V999.gCodSequencia
    lSelect = lSelect & " order by Cod_Sequencia desc "

    If gM46V999.gfObtRegistro(lbdTbm_Coef_PI, lSelect, lrsTbm_Coef_PI, lMensagem) = False Then
        Call gM46V999.gpFecha2(lbdTbm_Coef_PI)
        Call gM46V999.gpGraLog(2, lMensagem)
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    '2 - Se encontrar permite cadastro de Período Indenitário
    If Not lrsTbm_Coef_PI.EOF Then
        mfBuscaPerIndenitario = True
    End If
    
    'Supermercado SMART
    mfBuscaPerIndenitario = mfBuscaPerIndenitario Or _
                            (colCobert(intPosicao).ObjPai.Cod_Produto >= 1042) And _
                            (colCobert(intPosicao).ObjPai.Cod_Ramo = e_CodRamo_Empresarial) And _
                            (colCobert(intPosicao).ObjPai.ObjPai.Cod_Estip >= 2178 And colCobert(intPosicao).ObjPai.ObjPai.Cod_Estip <= 2182) And _
                            (colCobert(intPosicao).Cod_Cobert = 39 Or colCobert(intPosicao).Cod_Cobert = 3)
    
    Call gM46V999.gpFecha2(lbdTbm_Coef_PI)
End Function
