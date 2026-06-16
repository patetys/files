VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmT46V176A 
   AutoRedraw      =   -1  'True
   Caption         =   "Solicitação de inspeção"
   ClientHeight    =   5265
   ClientLeft      =   1650
   ClientTop       =   2010
   ClientWidth     =   9945
   Icon            =   "T46V176A.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5265
   ScaleWidth      =   9945
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   6000
      MaxLength       =   6
      TabIndex        =   9
      Top             =   1080
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox picUnchecked 
      Height          =   255
      Left            =   660
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   8
      Top             =   1470
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picChecked 
      Height          =   255
      Left            =   300
      Picture         =   "T46V176A.frx":000C
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1470
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picRadioUnChecked 
      Height          =   255
      Left            =   1380
      Picture         =   "T46V176A.frx":0106
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   6
      Top             =   1470
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picRadioChecked 
      Height          =   255
      Left            =   1020
      Picture         =   "T46V176A.frx":0648
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   5
      Top             =   1470
      Visible         =   0   'False
      Width           =   285
   End
   Begin MSFlexGridLib.MSFlexGrid flxSolicInspecao 
      Height          =   3270
      Left            =   30
      TabIndex        =   4
      Top             =   540
      Width           =   9870
      _ExtentX        =   17410
      _ExtentY        =   5768
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      AllowUserResizing=   1
   End
   Begin Threed.SSCommand cmdConfirmar 
      Height          =   300
      Left            =   3240
      TabIndex        =   0
      Top             =   4290
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Confirmar"
      ForeColor       =   8388608
   End
   Begin Threed.SSCommand cmdSair 
      Height          =   300
      Left            =   4845
      TabIndex        =   1
      Top             =   4290
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Sair"
      ForeColor       =   8388608
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1560
      Top             =   3810
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
            Picture         =   "T46V176A.frx":0B8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V176A.frx":0C9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V176A.frx":0DAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V176A.frx":1300
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LstvSolic 
      Height          =   3135
      Left            =   120
      TabIndex        =   3
      Top             =   825
      Visible         =   0   'False
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   5530
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label lblInformativo 
      Caption         =   "Selecione abaixo quais itens deverão ser inspecionados:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   75
      TabIndex        =   2
      Top             =   90
      Width           =   6255
   End
End
Attribute VB_Name = "frmT46V176A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mobjstcPed As stcA46V702B
Dim mobjstcPedLoc As stcA46V708B

Private Enum IconValues
    iconUnchecked = 1
    iconChecked = 2
    iconRadioChecked = 3
    iconRadioUnChecked = 4
End Enum

Private Sub Configurar_Form()
    Call gM46V999.gpCenForm(Me)

    With flxSolicInspecao
        'a) Quantidade inicial de colunas e linhas.
        .Cols = 10
        .Rows = 1
                       
        .Col = 0
        .TextMatrix(0, 0) = "Item"
        .ColWidth(0) = 400
        .RowHeight(0) = 500
        .CellAlignment = flexAlignCenterCenter
                
        .Col = 1
        .TextMatrix(0, 1) = "Conforme Normas"
        .CellAlignment = flexAlignCenterTop
        Set .CellPicture = picRadioUnChecked
        .CellPictureAlignment = flexAlignCenterBottom
        .ColWidth(1) = 1600
        
        .Col = 2
        .TextMatrix(0, 2) = "Dispensada"
        .CellAlignment = flexAlignCenterTop
        Set .CellPicture = picRadioUnChecked
        .CellPictureAlignment = flexAlignCenterBottom
        .ColWidth(2) = 1200
        
        .Col = 3
        .TextMatrix(0, 3) = "Obrigatória"
        .CellAlignment = flexAlignCenterTop
        Set .CellPicture = picRadioUnChecked
        .CellPictureAlignment = flexAlignCenterBottom
        .ColWidth(3) = 1100
        
        .Col = 4
        .TextMatrix(0, 4) = "Programada"
        .CellAlignment = flexAlignCenterTop
        Set .CellPicture = picRadioUnChecked
        .CellPictureAlignment = flexAlignCenterBottom
        .ColWidth(4) = 1100
        
        .TextMatrix(0, 5) = "Inspeção"
        .ColWidth(5) = 800
        .TextMatrix(0, 6) = "Local de Risco"
        .ColWidth(6) = flxSolicInspecao.Width / 5
        .TextMatrix(0, 7) = "Complemento"
        .ColWidth(7) = 1100
        .TextMatrix(0, 8) = "Valor de Risco"
        .ColWidth(8) = flxSolicInspecao.Width / 5
        .TextMatrix(0, 9) = "Já Solicitado"
        .ColWidth(9) = 1000
        .Col = 0
    End With
End Sub

Public Sub Carregar_Form(ByVal pobjstcPed As stcA46V702B)
    Set mobjstcPed = pobjstcPed
    
    Call Configurar_Form
    LstvSolic.SmallIcons = ImageList1
    LstvSolic.View = lvwList
    LstvSolic.LabelWrap = False
    LstvSolic.ColumnHeaders.Add , , "Item", LstvSolic.Width / 7
    LstvSolic.ColumnHeaders.Add , , "Local de risco", LstvSolic.Width / 4
    LstvSolic.ColumnHeaders.Add , , "Complemento", LstvSolic.Width / 5
    LstvSolic.ColumnHeaders.Add , , "Valor de Risco", LstvSolic.Width / 5
    LstvSolic.ColumnHeaders.Add , , "Já solicitado", LstvSolic.Width / 6
    ' Set View property to Report.
    LstvSolic.View = lvwReport

    With flxSolicInspecao
        .Rows = 1
        .FocusRect = flexFocusLight
        For Each mobjstcPedLoc In mobjstcPed.ColecaoPedLoc
            .Rows = .Rows + 1
            
            Call Carregar_Inspecao_Linha(.Rows - 1, mobjstcPedLoc.Num_Item, mobjstcPedLoc.Ind_StatusInspecao, _
            mobjstcPedLoc.Num_Insp, mobjstcPedLoc.Nom_Logr_Risco, mobjstcPedLoc.objstcPedInsp.Num_Logr_Risco, _
            mobjstcPedLoc.Dsc_Compl_Risco, mobjstcPedLoc.Val_Vr, mobjstcPedLoc.NOSSO_NUMERO)
        Next
    End With
    LstvSolic.Enabled = (pobjstcPed.Dat_Pri_Calculo = 0)
End Sub

Public Sub Carregar_Inspecao_Linha(ByVal lrow As Long, ByVal pNum_Item As Long, ByVal pInd_StatusInspecao As Integer, _
                                   ByVal pNum_Insp As Long, ByVal pNom_Logr_Risco As String, ByVal pNum_Logr_Risco As String, _
                                   ByVal pDsc_Compl_Risco As String, ByVal pVal_Vr As Double, ByVal pNosso_Numero As String)
    
    Dim lSql As String
    Dim lrst As clsYasRecordSet
    Dim lBD As ADODB.Connection
    Dim lMensagem As String
    Dim lI As Integer
    
    On Error GoTo Erro
    
    With flxSolicInspecao
        If lrow = 0 Then
            lrow = GetRow_ByNumItem(pNum_Item)
        End If
        
        .Row = lrow
        .Col = 0
        Set .CellPicture = picUnchecked
        .TextMatrix(.Row, 0) = pNum_Item
        .CellAlignment = flexAlignCenterCenter
        
        If gM46V111.gCod_plano > 1000 Then
            .Col = 1
            Set .CellPicture = picRadioChecked
            .CellPictureAlignment = flexAlignCenterCenter
            
        Else
            .Col = 1
            'Luiz Alberto - 04/11/2015
            If pInd_StatusInspecao = 0 Then
                Set .CellPicture = picRadioChecked
            Else
                Set .CellPicture = picRadioUnChecked
            End If
            .CellPictureAlignment = flexAlignCenterCenter
            
            .Col = 2
            If pInd_StatusInspecao = 1 Then
                Set .CellPicture = picRadioChecked
            Else
                Set .CellPicture = picRadioUnChecked
            End If
            
            .CellPictureAlignment = flexAlignCenterCenter
            .Col = 3
            If pInd_StatusInspecao = 2 Then
                Set .CellPicture = picRadioChecked
            Else
                Set .CellPicture = picRadioUnChecked
            End If
            .CellPictureAlignment = flexAlignCenterCenter
            
            .Col = 4
            If pInd_StatusInspecao = 3 Then
                Set .CellPicture = picRadioChecked
                .TextMatrix(.Row, 5) = pNum_Insp
            Else
                Set .CellPicture = picRadioUnChecked
                .TextMatrix(.Row, 5) = ""
            End If
            .CellPictureAlignment = flexAlignCenterCenter
        End If
        
        '*Luiz Alberto em 02/02/2016
        '****************************
        .TextMatrix(.Row, 6) = pNom_Logr_Risco & "," & pNum_Logr_Risco
        .TextMatrix(.Row, 7) = pDsc_Compl_Risco
        .TextMatrix(.Row, 8) = "R$ " & Format(pVal_Vr, "standard")
        
        If Not gM46V999.gfAbrBasDados("RAMOSDIVERSOS", gM46V999.gTipDIVERSOS, gM46V999.gEndDIVERSOS, lBD, lMensagem) Then
            Call gM46V999.gpGraLog(1, lMensagem)
            Exit Sub
        End If
    
        lSql = "Select * from ramosdiversos.dbo.tab_trans_insp where " & _
             " nosso_numero = " & gM46V999.gfFormatarCamposSQL(pNosso_Numero) & _
             " and num_item =" & pNum_Item
        If Not gM46V999.gfObtRegistro(lBD, lSql, lrst, lMensagem) Then
            Call gM46V999.gpGraLog(1, lMensagem)
            Exit Sub
        End If
            
        If lrst.EOF Then
            .TextMatrix(.Row, 9) = "Não"
        Else
            If lrst("Num_Insp") <> 0 Then
                .TextMatrix(.Row, 9) = "Sim (" & lrst("Num_Insp") & ")"
            Else
                .TextMatrix(.Row, 9) = "Sim "
            End If
        End If
    End With
    
    Exit Sub
Erro:
    Call gM46V999.gpGraLog(0, Me.Name + "_Carregar_Inspecao_Linha ERRO - " & Err & " " & Error)
End Sub

Private Sub cmdConfirmar_Click()
    Dim lcol As Integer
    Dim pMensagem As String
    Dim lTemItens As Boolean
    Dim lobjPedObs As clsA46V714A
    Dim lobjPedLoc As clsA46V708A
    Dim objTabPedLoc As stcA46V708B
    Dim objConexao As ADODB.Connection
    
    On Error GoTo Erro

    lTemItens = False
    For Each mobjstcPedLoc In mobjstcPed.ColecaoPedLoc
        flxSolicInspecao.Row = GetRow_ByNumItem(mobjstcPedLoc.Num_Item)
        
        flxSolicInspecao.Col = 1
        mobjstcPedLoc.Num_Insp = 0
        If flxSolicInspecao.CellPicture = picRadioChecked Then
            mobjstcPedLoc.Ind_StatusInspecao = e_StatusInspecao_ConfNormas
        Else
            flxSolicInspecao.Col = 2
            If flxSolicInspecao.CellPicture = picRadioChecked Then
                mobjstcPedLoc.Ind_StatusInspecao = e_StatusInspecao_Dispensada
            Else
                flxSolicInspecao.Col = 3
                If flxSolicInspecao.CellPicture = picRadioChecked Then
                    mobjstcPedLoc.Ind_StatusInspecao = e_StatusInspecao_obrigatoria
                Else
                    mobjstcPedLoc.Ind_StatusInspecao = e_StatusInspecao_programada
                    
                    If Trim(flxSolicInspecao.TextMatrix(flxSolicInspecao.Row, 5)) = "" Then
                        Call gM46V999.gpGraLog(1, "FGE0096 - O N.º/Inspeção deve estar preenchido.")
                        Exit Sub
                    End If
                    
                    mobjstcPedLoc.Num_Insp = flxSolicInspecao.TextMatrix(flxSolicInspecao.Row, 5)
                End If
            End If
        End If
        
        If Val(frmT46V101A.txtNumItem.Text) = mobjstcPedLoc.Num_Item Then
            Dim gCarga_Ant As Boolean
            
            gCarga_Ant = gCarga
            gCarga = True
            
            frmT46V122A.txtObservacao.Text = frmT46V101A.Nova_Observacao(mobjstcPedLoc.Ind_StatusInspecao, mobjstcPedLoc.Num_Insp, mobjstcPedLoc, frmT46V122A.txtObservacao.Text)
            mobjstcPedLoc.objstcPedObs.Observacao = Replace(frmT46V122A.txtObservacao.Text, vbCrLf, "&%")
            
            Select Case mobjstcPedLoc.Ind_StatusInspecao
                Case e_StatusInspecao_ConfNormas
                    frmT46V101A.optCondicaoInspecao(0).Value = True
                Case e_StatusInspecao_Dispensada
                    frmT46V101A.optCondicaoInspecao(1).Value = True
                Case e_StatusInspecao_obrigatoria
                    frmT46V101A.optCondicaoInspecao(2).Value = True
                Case e_StatusInspecao_programada
                    frmT46V101A.optCondicaoInspecao(3).Value = True
            End Select
            frmT46V101A.txtNumInsp.Text = mobjstcPedLoc.Num_Insp
            gCarga = gCarga_Ant
        Else
            If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, objConexao, pMensagem) = False Then
                Call gM46V999.gpGraLog(2, pMensagem)
                Exit Sub
            End If
            
            If mobjstcPedLoc.Ind_StatusInspecao = e_StatusInspecao_programada Then
                If Not gM46V999.gfValidaInspecao(mobjstcPedLoc.Num_Item, CDbl(flxSolicInspecao.TextMatrix(flxSolicInspecao.Row, 5)), mobjstcPedLoc.ObjPai.Num_Cgc_Cpf, mobjstcPedLoc.Num_Cep_Risco, mobjstcPedLoc.objstcPedInsp.Num_Logr_Risco, pMensagem, objConexao) Then
                    Call gM46V999.gpGraLog(1, pMensagem)
                    Exit Sub
                End If
            End If
            mobjstcPedLoc.objstcPedObs.Observacao = Replace(frmT46V101A.Nova_Observacao(mobjstcPedLoc.Ind_StatusInspecao, mobjstcPedLoc.Num_Insp, mobjstcPedLoc), vbCrLf, "&%")
        End If
       
        
    Next mobjstcPedLoc
    
    Call gM46V999.gpGraLog(1, "Operação realizada com sucesso!", vbInformation)
    Call cmdSair_Click
    Me.Refresh
    
    Exit Sub
Erro:
    Call gM46V999.gpGraLog(1, "frm46V176A.cmdConfirmar_Click" & Err.Number & " - " & Err.Description)
End Sub

Private Sub cmdSair_Click()
    Me.Visible = False
    frmT46V101A.SetFocus
End Sub

Private Sub flxSolicInspecao_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Integer
    Call Text1_LostFocus
    flxSolicInspecao.ScrollBars = flexScrollBarBoth

    'MsgBox x & " - " & y
    'If y > 2000 Then Exit Sub

    Select Case flxSolicInspecao.ColSel

        Case 1
        
            flxSolicInspecao.Col = 1
            
            If y < 400 Then
                flxSolicInspecao.Row = 0
                If flxSolicInspecao.CellPicture = picRadioChecked Then
                    For i = 0 To flxSolicInspecao.Rows - 1
                        flxSolicInspecao.Row = i
                        Set flxSolicInspecao.CellPicture = picRadioUnChecked
                    Next i
                Else
                    For i = 0 To flxSolicInspecao.Rows - 1
                        flxSolicInspecao.Row = i
                        Set flxSolicInspecao.CellPicture = picRadioChecked
                    Next i
                End If

                Call LimpaRadioCol(flxSolicInspecao.Col)
                Call LimpaTextBoxCol(5)

            Else

                If flxSolicInspecao.CellPicture = picRadioChecked Then
                    Set flxSolicInspecao.CellPicture = picRadioUnChecked
                Else
                    Set flxSolicInspecao.CellPicture = picRadioChecked
            End If
                
                Call LimpaRadioLin(flxSolicInspecao.Col)
                Call LimpaTextBoxLin(5)
                                
            End If
            
        Case 2
        
            flxSolicInspecao.Col = 2
        
            If y < 400 Then
                flxSolicInspecao.Row = 0
                If flxSolicInspecao.CellPicture = picRadioChecked Then
                    For i = 0 To flxSolicInspecao.Rows - 1
                        flxSolicInspecao.Row = i
                        Set flxSolicInspecao.CellPicture = picRadioUnChecked
                    Next i
                Else
                    For i = 0 To flxSolicInspecao.Rows - 1
                        flxSolicInspecao.Row = i
                        Set flxSolicInspecao.CellPicture = picRadioChecked
                    Next i
                End If
    
                Call LimpaRadioCol(flxSolicInspecao.Col)
                Call LimpaTextBoxCol(5)
                                
            Else
                If flxSolicInspecao.CellPicture = picRadioChecked Then
                    Set flxSolicInspecao.CellPicture = picRadioUnChecked
                Else
                    Set flxSolicInspecao.CellPicture = picRadioChecked
                End If
                
                Call LimpaRadioLin(flxSolicInspecao.Col)
                Call LimpaTextBoxLin(5)
                
            End If
        
        Case 3
        
            flxSolicInspecao.Col = 3
            
            If y < 400 Then
                flxSolicInspecao.Row = 0
                If flxSolicInspecao.CellPicture = picRadioChecked Then
                    For i = 0 To flxSolicInspecao.Rows - 1
                        flxSolicInspecao.Row = i
                        Set flxSolicInspecao.CellPicture = picRadioUnChecked
                    Next i
                Else
                    For i = 0 To flxSolicInspecao.Rows - 1
                        flxSolicInspecao.Row = i
                        Set flxSolicInspecao.CellPicture = picRadioChecked
                    Next i
                End If
                
                Call LimpaRadioCol(flxSolicInspecao.Col)
                Call LimpaTextBoxCol(5)
                
            Else
                If flxSolicInspecao.CellPicture = picRadioChecked Then
                    Set flxSolicInspecao.CellPicture = picRadioUnChecked
                Else
                    Set flxSolicInspecao.CellPicture = picRadioChecked
                End If
                
                Call LimpaRadioLin(flxSolicInspecao.Col)
                Call LimpaTextBoxLin(5)
                
            End If
            
            'Luiz Alberto - 05/11/2015 - anterior
            'Call LimpaRadioCol(flxSolicInspecao.Col)
        
        Case 4
        
            flxSolicInspecao.Col = 4
            
            If y < 400 Then
                flxSolicInspecao.Row = 0
                If flxSolicInspecao.CellPicture = picRadioChecked Then
                    For i = 0 To flxSolicInspecao.Rows - 1
                        flxSolicInspecao.Row = i
                        Set flxSolicInspecao.CellPicture = picRadioUnChecked
                    Next i
                Else
                    For i = 0 To flxSolicInspecao.Rows - 1
                        flxSolicInspecao.Row = i
                        Set flxSolicInspecao.CellPicture = picRadioChecked
                    Next i
                End If
                
                Call LimpaRadioCol(flxSolicInspecao.Col)
                
            Else
            
                If flxSolicInspecao.CellPicture = picRadioChecked Then
                    Set flxSolicInspecao.CellPicture = picRadioUnChecked
                Else
                    Set flxSolicInspecao.CellPicture = picRadioChecked
                End If
                
                Call LimpaRadioLin(flxSolicInspecao.Col)
                
            End If
            
            'Luiz Alberto - 05/11/2015 - anterior
            'Call LimpaRadioCol(flxSolicInspecao.Col)
        Case 5
            
            flxSolicInspecao.Col = 4
            
            If flxSolicInspecao.CellPicture = picRadioChecked Then
                flxSolicInspecao.Col = 5
                Text1.Move flxSolicInspecao.Left + flxSolicInspecao.CellLeft, _
                           flxSolicInspecao.Top + flxSolicInspecao.CellTop, _
                           flxSolicInspecao.CellWidth, flxSolicInspecao.CellHeight
                Text1.Tag = flxSolicInspecao.Row & "," & flxSolicInspecao.Col
                
                If (flxSolicInspecao.TextMatrix(flxSolicInspecao.Row, flxSolicInspecao.Col)) <> Text1.Text Then
                   Text1.Text = flxSolicInspecao.TextMatrix(flxSolicInspecao.Row, flxSolicInspecao.Col)
                End If
                   
                Text1.Visible = True
            End If
            
        End Select
    
End Sub

Private Sub LimpaRadioCol(ByRef num_Col As Integer)
    Dim i As Integer
    Dim j As Integer
    
    For i = 0 To flxSolicInspecao.Cols - 1
        flxSolicInspecao.Col = i
        If flxSolicInspecao.CellPicture = picRadioChecked Or flxSolicInspecao.CellPicture = picRadioUnChecked Then
            If flxSolicInspecao.Col <> num_Col Then
                For j = 0 To flxSolicInspecao.Rows - 1
                    flxSolicInspecao.Row = j
                    Set flxSolicInspecao.CellPicture = picRadioUnChecked
                Next j
            End If
        End If
    Next i
    
    flxSolicInspecao.Col = num_Col
End Sub

Private Sub LimpaRadioLin(ByRef num_Col As Integer)
    Dim i As Integer
    
    For i = 0 To flxSolicInspecao.Cols - 1
        flxSolicInspecao.Col = i
        If flxSolicInspecao.CellPicture = picRadioChecked Or flxSolicInspecao.CellPicture = picRadioUnChecked Then
            If flxSolicInspecao.Col <> num_Col Then
                Set flxSolicInspecao.CellPicture = picRadioUnChecked
                
            End If
        End If
    Next i

    flxSolicInspecao.Col = num_Col
End Sub
Private Sub LimpaTextBoxLin(ByRef num_Col As Integer)
    Dim i As Integer
    
    For i = 0 To flxSolicInspecao.Cols - 1
        flxSolicInspecao.Col = i
        If flxSolicInspecao.Col = num_Col Then
           Text1.Text = ""
           flxSolicInspecao.TextMatrix(flxSolicInspecao.Row, num_Col) = Text1.Text
        End If
    Next i
    
End Sub

Private Sub LimpaTextBoxCol(ByRef num_Col As Integer)
    
    Dim j As Integer
    
    flxSolicInspecao.Col = num_Col
    
    For j = 1 To flxSolicInspecao.Rows - 1
        flxSolicInspecao.Row = j
        Text1.Text = ""
        flxSolicInspecao.TextMatrix(flxSolicInspecao.Row, num_Col) = Text1.Text
    Next j
                
End Sub

Private Sub Form_Activate()
    flxSolicInspecao.Enabled = Not (gM46V999.gDesabFormsCotacao Or mobjstcPed.getPlano.EhPlaninho) And gM46V999.gTipoAcesso <> DIGITACAO
    flxSolicInspecao.Enabled = flxSolicInspecao.Enabled And Not (mobjstcPed.Dat_Pri_Calculo <> 0 And gM46V999.gTipoAcesso = CALCULO And gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda And (UCase(gM46V999.gSerDIVERSOS) Like "*MIX*" = False) And Not UCase(gM46V999.gNomUsuMainFrame) Like "DSIS*")
    cmdConfirmar.Enabled = flxSolicInspecao.Enabled
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        Me.Visible = False
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii

    Case 8    'BackSpace
        Exit Sub

    Case 13    'Enter
        Exit Sub
        
    Case 22    'Ctrl+V
        If flxSolicInspecao.Col = 5 Then
            If IsNumeric(Trim(Clipboard.GetText)) Then
                Text1.Text = Trim(Clipboard.GetText)
            End If
        End If
        
    Case 48 To 57    'Zero a nove
        If flxSolicInspecao.Col = 5 Then
            Text1.Text = Text1.Text & Chr(KeyAscii)
        End If
        
    Case Else
        Beep
    End Select

    KeyAscii = 0
    Text1.SelStart = Len(Text1.Text)
End Sub

Private Sub Text1_LostFocus()
    Dim lcol As Integer
    Dim lrow As Integer

    If Len(Text1.Tag) Then
        lrow = Left(Text1.Tag, InStr(Text1.Tag, ","))
        lcol = 5 'Right(Text1.Tag, InStr(Text1.Tag, ",") - 1)

        If lrow < flxSolicInspecao.Rows Then
            flxSolicInspecao.TextMatrix(lrow, lcol) = Text1.Text
        End If
    End If
    Text1.Visible = False
End Sub

Private Sub Form_Resize()

    On Error GoTo Erro
    
    If Me.Height > 1500 And Me.Width > 1500 Then
        With flxSolicInspecao
            .Left = 30
            .Top = 530
            .Width = Me.Width - 270
            .Height = Me.Height - 1600
        End With
        
        cmdConfirmar.Top = Me.Height - 1000
        cmdConfirmar.Left = CInt((Me.Width - (cmdConfirmar.Width + cmdSair.Width + 60)) / 2) - 270
        cmdSair.Top = cmdConfirmar.Top
        cmdSair.Left = cmdConfirmar.Left + cmdConfirmar.Width + 60
    End If
    
    Exit Sub
Erro:
    Call gM46V999.gpGraLog(0, Me.Name + "_Form_Resize ERRO - " & Err & " " & Error)
End Sub

Private Function GetRow_ByNumItem(ByVal pNum_Item As Long) As Integer
    Dim lrow As Integer
    
    On Error GoTo Erro
    
    lrow = 1
    Do While lrow < flxSolicInspecao.Rows
        If flxSolicInspecao.TextMatrix(lrow, 0) = pNum_Item Then
            GetRow_ByNumItem = lrow
            Exit Do
        End If
        lrow = lrow + 1
    Loop
    
    Exit Function
Erro:
    GetRow_ByNumItem = 0
    Call gM46V999.gpGraLog(0, Me.Name + "_GetRow_ByNumItem - " & Err & " " & Error)
End Function

