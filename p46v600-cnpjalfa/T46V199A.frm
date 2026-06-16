VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmT46V199A 
   Caption         =   "Gerenciador de Coberturas"
   ClientHeight    =   6690
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12210
   Icon            =   "T46V199A.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6690
   ScaleWidth      =   12210
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboNum_Item 
      Height          =   315
      ItemData        =   "T46V199A.frx":0A02
      Left            =   3480
      List            =   "T46V199A.frx":0A04
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   960
      Width           =   5595
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12210
      _ExtentX        =   21537
      _ExtentY        =   1164
      ButtonWidth     =   1296
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgLista"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Editar"
            Key             =   "EditarCobert"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Replicar"
            Key             =   "Replicar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair"
            Key             =   "Sair"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraReplicar 
      Caption         =   "Replicar nos demais itens:"
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   60
      TabIndex        =   1
      Top             =   750
      Width           =   3285
      Begin VB.CheckBox chkReplicar 
         Caption         =   "Val IS"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   2
         ToolTipText     =   "Valor da Importancia Segurada"
         Top             =   300
         Width           =   735
      End
      Begin VB.CheckBox chkReplicar 
         Caption         =   "PI Meses"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   2130
         TabIndex        =   4
         ToolTipText     =   "Periodo Indenitário"
         Top             =   300
         Width           =   975
      End
      Begin VB.CheckBox chkReplicar 
         Caption         =   "Franquia"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   990
         TabIndex        =   3
         ToolTipText     =   "Dados de Franquia (Percentual, Val Minimo, Desconto)"
         Top             =   300
         Width           =   915
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   345
      Left            =   30
      TabIndex        =   8
      Top             =   6300
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   7
      Top             =   6255
      Width           =   12210
      _ExtentX        =   21537
      _ExtentY        =   767
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid flxCoberturas 
      Height          =   4005
      Left            =   60
      TabIndex        =   6
      Top             =   1470
      Width           =   12075
      _ExtentX        =   21299
      _ExtentY        =   7064
      _Version        =   393216
      Rows            =   5
      Cols            =   4
      FixedCols       =   0
      BackColor       =   -2147483637
      ForeColor       =   8388608
      BackColorFixed  =   -2147483636
      ForeColorFixed  =   16777215
      ForeColorSel    =   12648447
      BackColorBkg    =   -2147483633
      GridColorFixed  =   4210752
      AllowBigSelection=   0   'False
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin MSComctlLib.ImageList imgLista 
      Left            =   11250
      Top             =   540
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
            Picture         =   "T46V199A.frx":0A06
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V199A.frx":0D58
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V199A.frx":10AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V199A.frx":13FC
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmT46V199A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private objstcA46V708B As stcA46V708B

Public Enum T46V199_Grid
    Codigo = 1
    Descricao = 2
    LimMaxInden = 3
    PerFranq = 4
    Unidade = 5
    FranqMin = 6
    'Multiplo = 7
    Desconto = 7
    PIndenitario = 8
End Enum

Public Sub CarregarForm(ByVal pobjstcA46V708B As stcA46V708B)
    Set objstcA46V708B = pobjstcA46V708B
    Call ConfigGrid
  
    Dim Item As stcA46V708B
    With cboNum_Item
        .Clear
        For Each Item In objstcA46V708B.ObjPai.ColecaoPedLoc
            Call .AddItem(Format(Item.Num_Item, "0000") & " - " & Item.gNom_Class)
            '.ItemData(.NewIndex) = item.Num_Item
        Next Item
        
        Dim lI As Integer
        For lI = 0 To (.ListCount - 1)
            If objstcA46V708B.Num_Item = CInt(Left(.List(lI), 4)) Then
                .ListIndex = lI
                Exit For
            End If
        Next lI
        
        
    End With
    
    Me.Show vbModal
End Sub

Private Sub ConfigGrid()
    flxCoberturas.Rows = 1
    flxCoberturas.Cols = 9
    flxCoberturas.Enabled = True
    
    flxCoberturas.TextMatrix(0, 0) = ""
    flxCoberturas.ColWidth(0) = 250
    
    flxCoberturas.TextMatrix(0, T46V199_Grid.Codigo) = "Cód"
    flxCoberturas.ColAlignment(T46V199_Grid.Codigo) = flexAlignCenterCenter
    flxCoberturas.ColWidth(1) = 470
    
    flxCoberturas.TextMatrix(0, T46V199_Grid.Descricao) = "Cobertura"
    flxCoberturas.ColAlignment(T46V199_Grid.Descricao) = flexAlignCenterCenter
    flxCoberturas.ColWidth(2) = 4750
    
    flxCoberturas.TextMatrix(0, T46V199_Grid.LimMaxInden) = "Lim Max Indenização"
    flxCoberturas.ColAlignment(T46V199_Grid.LimMaxInden) = flexAlignCenterCenter
    flxCoberturas.ColWidth(T46V199_Grid.LimMaxInden) = 1575
    
    flxCoberturas.TextMatrix(0, T46V199_Grid.PerFranq) = "%Fraq"
    flxCoberturas.ColAlignment(T46V199_Grid.PerFranq) = flexAlignCenterCenter
    flxCoberturas.ColWidth(T46V199_Grid.PerFranq) = 615
    
    flxCoberturas.TextMatrix(0, T46V199_Grid.Unidade) = "#"
    flxCoberturas.ColAlignment(T46V199_Grid.Unidade) = flexAlignCenterCenter
    flxCoberturas.ColWidth(6) = 555
    
    flxCoberturas.TextMatrix(0, T46V199_Grid.FranqMin) = "Franquia Min"
    flxCoberturas.ColAlignment(T46V199_Grid.FranqMin) = flexAlignCenterCenter
    flxCoberturas.ColWidth(T46V199_Grid.FranqMin) = 1050
    
'    flxCoberturas.TextMatrix(0, T46V199_Grid.Multiplo) = "Múltiplo"
'    flxCoberturas.ColAlignment(T46V199_Grid.Multiplo) = flexAlignCenterCenter
'    flxCoberturas.ColWidth(T46V199_Grid.Multiplo) = 0 '855
    
    flxCoberturas.TextMatrix(0, T46V199_Grid.Desconto) = "%Desc"
    flxCoberturas.ColAlignment(T46V199_Grid.Desconto) = flexAlignCenterCenter
    flxCoberturas.ColWidth(T46V199_Grid.Desconto) = 600
    
    flxCoberturas.TextMatrix(0, T46V199_Grid.PIndenitario) = "Pi"
    flxCoberturas.ColAlignment(T46V199_Grid.PIndenitario) = flexAlignCenterCenter
    flxCoberturas.ColWidth(T46V199_Grid.PIndenitario) = 510
End Sub

Private Sub Carrega_Grid(pobjstcA46V708B As stcA46V708B)
    Dim objstcA46V704B As stcA46V704B
    
        With flxCoberturas
        .Rows = 1
        Do While .Rows <= pobjstcA46V708B.ColecaoPedCobert.Count
            .Redraw = False
            .Rows = .Rows + 1
            
            Set objstcA46V704B = pobjstcA46V708B.ColecaoPedCobert(.Rows - 1)
            .TextMatrix(.Rows - 1, T46V199_Grid.Codigo) = objstcA46V704B.Cod_Cobert
            .Row = .Rows - 1
            .Col = T46V199_Grid.Codigo
            .CellAlignment = flexAlignRightCenter
            
            .TextMatrix(.Rows - 1, T46V199_Grid.Descricao) = objstcA46V704B.gNomeCobertura
            .Row = .Rows - 1
            .Col = T46V199_Grid.Descricao
            .CellAlignment = flexAlignLeftCenter
            
            .TextMatrix(.Rows - 1, T46V199_Grid.LimMaxInden) = Format(objstcA46V704B.Val_IS, "###,###,##0.00")
            .Row = .Rows - 1
            .Col = T46V199_Grid.LimMaxInden
            .CellAlignment = flexAlignRightCenter
            
            .TextMatrix(.Rows - 1, T46V199_Grid.PerFranq) = objstcA46V704B.Per_Franq
            .TextMatrix(.Rows - 1, T46V199_Grid.Unidade) = IIf(objstcA46V704B.Num_Meses_PI > 0, "dias", "R$")
            
            .TextMatrix(.Rows - 1, T46V199_Grid.FranqMin) = Format(objstcA46V704B.Val_Franq_Min, "###,###,##0.00")
            .Row = .Rows - 1
            .Col = T46V199_Grid.FranqMin
            .CellAlignment = flexAlignRightCenter
            
            '.TextMatrix(.Rows - 1, T46V199_Grid.Multiplo) = objstcA46V704B.gMultiploFraq & "x"
            
            .TextMatrix(.Rows - 1, T46V199_Grid.Desconto) = Format(objstcA46V704B.Per_Desc_Franq, "#0.00")
            .Row = .Rows - 1
            .Col = T46V199_Grid.Desconto
            .CellAlignment = flexAlignRightCenter
            
            .TextMatrix(.Rows - 1, T46V199_Grid.PIndenitario) = objstcA46V704B.Num_Meses_PI
            
            .Row = 1
            .Col = 0
            .ColSel = 8
            .Redraw = True
        Loop
        End With
    
End Sub

Private Sub Replicar()
    Dim Item As stcA46V708B
    Dim Cobert As stcA46V704B
    Dim CobertBase As stcA46V704B
    Dim lQtdln As Integer
    Dim blnSelecionada As Boolean
    Dim lObjPedCobert As clsA46V704A
    Dim lObjPedCobertFranquia As clsA46V738A
    Dim lMsgErro As String
    Dim lTemObs As String
    Dim i As Integer
    Dim lMsg As String
    
    On Error GoTo Erro
    Screen.MousePointer = vbHourglass
    LockWindowUpdate flxCoberturas.hWnd
    
    Set lObjPedCobert = New clsA46V704A
    Set lObjPedCobert.M46V999 = gM46V999
    Set lObjPedCobertFranquia = New clsA46V738A
    
    If chkReplicar(0).Value = Unchecked And chkReplicar(1).Value = Unchecked And chkReplicar(2).Value = Unchecked Then
        MsgBox "Selecione ao menos um dado a ser replicado aos demais itens.", vbExclamation
        GoTo Fim
    End If
    
    blnSelecionada = False
    i = 1
    Do While i < flxCoberturas.Rows
        flxCoberturas.Col = 0
        flxCoberturas.Row = i
        If flxCoberturas.CellPicture <> 0 Then
            blnSelecionada = True
            Exit Do
        End If
        i = i + 1
    Loop
    
    If blnSelecionada = 0 Then
        MsgBox "Selecione ao menos uma cobertura para ser replicada aos demais itens.", vbExclamation
        GoTo Fim
    End If
    
    lMsg = "Confirma a replicação dos seguintes dados: " & vbCrLf
    lMsg = lMsg & IIf(chkReplicar(0).Value, " - " & chkReplicar(0).ToolTipText & vbCrLf, Empty)
    lMsg = lMsg & IIf(chkReplicar(1).Value, " - " & chkReplicar(1).ToolTipText & vbCrLf, Empty)
    lMsg = lMsg & IIf(chkReplicar(2).Value, " - " & chkReplicar(2).ToolTipText & vbCrLf, Empty)
    lMsg = lMsg & "Para todos os demais itens desta cotação ?"
    
    If MsgBox(lMsg, vbQuestion + vbYesNo) = vbYes Then
        For Each Item In objstcA46V708B.ObjPai.ColecaoPedLoc
            'If Item.Num_Item <> objstcA46V708B.Num_Item Then
                For Each Cobert In Item.ColecaoPedCobert
                    
                    i = 1
                    blnSelecionada = False
                    Do While i < flxCoberturas.Rows
                        If flxCoberturas.TextMatrix(i, T46V199_Grid.Codigo) = Cobert.Cod_Cobert Then
                            flxCoberturas.Col = 0
                            flxCoberturas.Row = i
                            If flxCoberturas.CellPicture <> 0 Then
                                blnSelecionada = True
                                Exit Do
                            End If
                        End If
                        i = i + 1
                    Loop
                    
                    If blnSelecionada Then
                        For Each CobertBase In objstcA46V708B.ColecaoPedCobert
                            If Cobert.Cod_Cobert = CobertBase.Cod_Cobert Then
                                If chkReplicar(0).Value = Checked Then
                                    Cobert.Val_IS = CobertBase.Val_IS
                                End If
                                
                                If chkReplicar(1).Value = Checked Then
                                    Cobert.AlterouFranquia = "SIM"
                                    Cobert.Per_Franq = CobertBase.Per_Franq
                                    Cobert.Val_Franq_Min = CobertBase.Val_Franq_Min
                                    Cobert.Per_Desc_Franq = CobertBase.Per_Desc_Franq
                                    Cobert.TextoFranquia = gM46V999.gfTextoFranquia(gM46V999.CDblx(Cobert.Val_Franq_Min), gM46V999.CDblx(Cobert.Per_Franq), "", Cobert.Cod_Cobert, gM46V999.CDblx(Cobert.Num_Meses_PI))
                                    
                                    If Cobert.ColecaoNovaFranquia.Count > 0 Then
                                        Cobert.ColecaoNovaFranquia(1).Texto_Franquia = Cobert.TextoFranquia
                                        Cobert.ColecaoNovaFranquia(1).Multiplo_Franquia = IIf(Cobert.Per_Desc_Franq = 0, "1", "3")
                                        Cobert.ColecaoNovaFranquia(1).Per_Franq = Cobert.Per_Franq
                                        Cobert.ColecaoNovaFranquia(1).Val_Min_Franq = Cobert.Val_Franq_Min
                                        
                                        If CobertBase.ColecaoNovaFranquia.Count > 0 Then
                                            Cobert.ColecaoNovaFranquia(1).Texto_Observacoes = CobertBase.ColecaoNovaFranquia(1).Texto_Observacoes
                                            lTemObs = IIf(Len(Cobert.ColecaoNovaFranquia(1).Texto_Observacoes) > 0, "*", "")
                                            Cobert.TextoFranquia = gM46V999.gfTextoFranquia(gM46V999.CDblx(Cobert.Val_Franq_Min), gM46V999.CDblx(Cobert.Per_Franq), lTemObs, Cobert.Cod_Cobert, gM46V999.CDblx(Cobert.Num_Meses_PI))
                                        End If
                                    End If
                                End If
                                
                                If chkReplicar(2).Value = Checked Then
                                    Cobert.Num_Meses_PI = CobertBase.Num_Meses_PI
                                End If
                                
                                If Not lObjPedCobert.mfExcluir(Cobert, "", "", False, Cobert.Cod_Cobert) Then
                                    lMsgErro = lObjPedCobert.UltimoErro
                                    GoTo Erro
                                End If

                                If Not lObjPedCobert.mfIncluir(Cobert, "", "", False, False, "") Then
                                    lMsgErro = lObjPedCobert.UltimoErro
                                    GoTo Erro
                                End If
                                
                                If Cobert.ColecaoNovaFranquia.Count > 0 Then
                                    Set lObjPedCobertFranquia.M46V999 = gM46V999
                                    If Not lObjPedCobertFranquia.mfExcluir(Cobert.ColecaoNovaFranquia(1), "", "", Cobert.Cod_Cobert) Then
                                        lMsgErro = lObjPedCobert.UltimoErro
                                        GoTo Erro
                                    End If
                                    
                                    If Not lObjPedCobertFranquia.mfIncluir(Cobert.ColecaoNovaFranquia(1), "", "", "") Then
                                        lMsgErro = lObjPedCobert.UltimoErro
                                        GoTo Erro
                                    End If
                                End If
                                
                                Exit For
                            End If
                        Next CobertBase
                    End If
                    
                Next Cobert
            'End If
        Next Item
                
        Unload Me
    End If
    
    GoTo Fim
    Resume 0
Erro:
    If Err.Number <> 0 Then
        lMsgErro = lMsgErro & " " & Err.Number & " - " & Err.Description
    End If
    MsgBox lMsgErro
Fim:
    Screen.MousePointer = vbDefault
    LockWindowUpdate False
End Sub

Private Sub Salvar_Item()
    Dim Cobert As stcA46V704B
    Dim lObjPedCobert As clsA46V704A
    Dim lObjPedCobertFranquia As clsA46V738A
    Dim lMsgErro As String
    Dim lTemObs As String
    
    On Error GoTo Erro
    Screen.MousePointer = vbHourglass
    LockWindowUpdate flxCoberturas.hWnd
    
    Set lObjPedCobert = New clsA46V704A
    Set lObjPedCobert.M46V999 = gM46V999
    Set lObjPedCobertFranquia = New clsA46V738A
    Set lObjPedCobertFranquia.M46V999 = gM46V999
    
    For Each Cobert In objstcA46V708B.ColecaoPedCobert
    
        If Cobert.ColecaoNovaFranquia.Count > 0 Then
            lTemObs = IIf(Len(Cobert.ColecaoNovaFranquia(1).Texto_Observacoes) > 0, "*", "")
            Cobert.TextoFranquia = gM46V999.gfTextoFranquia(gM46V999.CDblx(Cobert.Val_Franq_Min), gM46V999.CDblx(Cobert.Per_Franq), lTemObs, Cobert.Cod_Cobert, gM46V999.CDblx(Cobert.Num_Meses_PI))
            Cobert.ColecaoNovaFranquia(1).Texto_Franquia = Cobert.TextoFranquia
            Cobert.ColecaoNovaFranquia(1).Multiplo_Franquia = IIf(Cobert.Per_Desc_Franq = 0, "1", "3")
            Cobert.ColecaoNovaFranquia(1).Per_Franq = Cobert.Per_Franq
            Cobert.ColecaoNovaFranquia(1).Val_Min_Franq = Cobert.Val_Franq_Min
        Else
            Cobert.TextoFranquia = gM46V999.gfTextoFranquia(gM46V999.CDblx(Cobert.Val_Franq_Min), gM46V999.CDblx(Cobert.Per_Franq), "", Cobert.Cod_Cobert, gM46V999.CDblx(Cobert.Num_Meses_PI))
        End If
        
        If Not lObjPedCobert.mfExcluir(Cobert, "", "", False, Cobert.Cod_Cobert) Then
            lMsgErro = lObjPedCobert.UltimoErro
            GoTo Erro
        End If
        
        If Not lObjPedCobert.mfIncluir(Cobert, "", "", False, False, "") Then
            lMsgErro = lObjPedCobert.UltimoErro
            GoTo Erro
        End If
        
        If Cobert.ColecaoNovaFranquia.Count > 0 Then
            
            If Not lObjPedCobertFranquia.mfExcluir(Cobert.ColecaoNovaFranquia(1), "", "", Cobert.Cod_Cobert) Then
                lMsgErro = lObjPedCobert.UltimoErro
                GoTo Erro
            End If
            
            If Not lObjPedCobertFranquia.mfIncluir(Cobert.ColecaoNovaFranquia(1), "", "", "") Then
                lMsgErro = lObjPedCobert.UltimoErro
                GoTo Erro
            End If
        End If
    Next Cobert
    
    GoTo Fim:
Erro:
    If Err.Number <> 0 Then
        lMsgErro = lMsgErro & " " & Err.Number & " - " & Err.Description
    End If
    MsgBox lMsgErro
Fim:
    Screen.MousePointer = vbDefault
End Sub

Private Sub cboNum_Item_Click()
    Set objstcA46V708B = objstcA46V708B.ObjPai.gColecaoPedLoc_ByNum_Item(CInt(Left(cboNum_Item.Text, 4)))
    Call Carrega_Grid(objstcA46V708B)
End Sub

Private Sub flxCoberturas_DblClick()
    Call frmT46V200A.CarregarForm(objstcA46V708B.ColecaoPedCobert, flxCoberturas.Row)
    Call Carrega_Grid(objstcA46V708B)
End Sub

Private Sub flxCoberturas_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    With flxCoberturas
        If .MouseCol = 0 Then
            .Row = .MouseRow
            If .CellPicture = 0 Then
                Set .CellPicture = imgLista.ListImages.Item(4).Picture 'picChecked.Picture
                .CellPictureAlignment = flexAlignCenterCenter
            Else
                Set .CellPicture = Nothing
            End If
        End If
    End With
End Sub

Private Sub Form_Load()
    Call ConfigGrid
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If UnloadMode = 0 Then
        Cancel = 1
        MsgBox "Utilize o botão sair!", vbInformation, Me.Caption
        Exit Sub
    End If

    If Not frmT46V101A.objstcPedido Is Nothing Then
        frmT46V101A.txtISBas.Text = 0
        Call frmT46V117A.mpSelecionarCoberturas(frmT46V101A.objstcPedLoc.ColecaoPedCobert, , True, False, False)
        Call frmT46V179A.mpSelecionarFranquias(frmT46V101A.objstcPedLoc.ColecaoPedCobert)
        Call gpLimPremios
    End If
End Sub

Private Sub Form_Resize()

    On Error GoTo Erro
    
    If Me.Height > 1500 And Me.Width > 1500 Then
        With flxCoberturas
            .Left = 30
            .Top = Toolbar.Height + fraReplicar.Height + 100
            .Width = Me.Width - 270
            .Height = Me.Height - .Top - 1300
        End With
        
        'cmdOk.Top = Me.Height - 1000
        'cmdOk.Left = CInt((Me.Width - (cmdOk.Width + cmdSalvar.Width + 60)) / 2) - 270
        'cmdSalvar.Top = cmdOk.Top
        'cmdSalvar.Left = cmdOk.Left + cmdOk.Width + 60
    End If
    
    Exit Sub
Erro:
    Call gM46V999.gpGraLog(0, Me.Name + "_Form_Resize ERRO - " & Err & " " & Error)
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button
        Case "Editar"
            Call frmT46V200A.CarregarForm(objstcA46V708B.ColecaoPedCobert, flxCoberturas.Row)
            Call Carrega_Grid(objstcA46V708B)
        
        Case "Replicar"
            Call Replicar
            
        Case "Sair"
            Salvar_Item
            Unload Me
    End Select
End Sub
