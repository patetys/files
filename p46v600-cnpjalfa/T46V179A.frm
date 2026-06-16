VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmT46V179A 
   Caption         =   "Textos de Franquia"
   ClientHeight    =   3030
   ClientLeft      =   5370
   ClientTop       =   3465
   ClientWidth     =   9960
   ControlBox      =   0   'False
   Icon            =   "T46V179A.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3030
   ScaleWidth      =   9960
   Begin VB.CommandButton cmdEditarObservacao 
      Caption         =   "Editar Obs."
      Height          =   420
      Left            =   90
      TabIndex        =   2
      Top             =   2505
      Width           =   1410
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   420
      Left            =   8475
      TabIndex        =   1
      Top             =   2505
      Width           =   1410
   End
   Begin MSFlexGridLib.MSFlexGrid flxCobertura 
      Height          =   2340
      Left            =   68
      TabIndex        =   0
      Top             =   90
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   4128
      _Version        =   393216
      Rows            =   5
      Cols            =   4
      FixedCols       =   0
      BackColor       =   -2147483644
      ForeColor       =   8388608
      BackColorFixed  =   -2147483636
      ForeColorFixed  =   16777215
      ForeColorSel    =   12648447
      BackColorBkg    =   -2147483633
      GridColorFixed  =   4210752
      AllowBigSelection=   0   'False
      GridLines       =   2
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "frmT46V179A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEditarObservacao_Click()
    frmT46V186A.Show 1
End Sub

Private Sub cmdOk_Click()
    Me.Hide
End Sub

Private Sub Form_Activate()
    If gM46V999.gCodProduto < 1080 Then
        cmdEditarObservacao.Visible = False
        flxCobertura.ColWidth(4) = 0
    Else
        cmdEditarObservacao.Visible = True
        flxCobertura.ColWidth(4) = 3500
    End If
End Sub

Private Sub Form_Load()
    '1. Posiciona form.
    Call gM46V999.gpCenForm(Me)
    Dim i As Integer

    With flxCobertura
        .Rows = 1
        .Cols = 5
        .TextMatrix(0, 0) = "Código"
        .TextMatrix(0, 1) = "Cobertura"
        .TextMatrix(0, 2) = "Franquia"
        .TextMatrix(0, 3) = "% Desc"
        .TextMatrix(0, 4) = "Observação"

        .ColWidth(0) = 600
        .ColWidth(1) = 3500
        .ColWidth(2) = 2000
        .ColWidth(3) = 700
        .ColWidth(4) = 3500

        .ColAlignment(2) = flexAlignLeftCenter
        
        If gM46V999.gCodProduto < 1080 Then
            cmdEditarObservacao.Visible = False
            .ColWidth(4) = 0
        End If
    End With
    
End Sub

Private Sub Form_Resize()
    If Me.Width < 10200 Then Me.Width = 10200
    If Me.Height < 3600 Then Me.Height = 3600
    
    flxCobertura.Width = Me.Width - 400
    flxCobertura.Height = Me.Height - 1200
    cmdEditarObservacao.Top = Me.Height - 1050
    cmdOK.Top = Me.Height - 1050
    cmdOK.Left = Me.Width - cmdOK.Width - 400
End Sub

Public Sub mpSelecionarFranquias(ByVal pColecaoPedCobert As Collection)
    Dim i As Integer
    
    With frmT46V117A.flxSelecionadas
        flxCobertura.Rows = .Rows
        For i = 1 To .Rows - 1
            If .TextMatrix(i, ColCodCobert) <> Empty And .TextMatrix(i, ColValLMG) > 0 Then
                flxCobertura.TextMatrix(i, 0) = .TextMatrix(i, ColCodCobert)
                flxCobertura.TextMatrix(i, 1) = .TextMatrix(i, ColDescFranquia)
                If Trim(flxCobertura.TextMatrix(i, 1)) = "" Then
                    flxCobertura.TextMatrix(i, 1) = .TextMatrix(i, 1)
                End If
                flxCobertura.TextMatrix(i, 2) = Trim(.TextMatrix(i, ColTextoFranquia))
                flxCobertura.TextMatrix(i, 3) = Trim(.TextMatrix(i, ColDesconto))
                flxCobertura.TextMatrix(i, 4) = mpGetObservacaoFranquia(.TextMatrix(i, ColCodCobert), pColecaoPedCobert)
            End If
        Next
    End With
    
End Sub

Public Sub mpAtualizarFranquias(ByVal pColecaoPedCobert As Collection)
    Dim i As Integer
    Dim j As Integer
    Dim lRecriarGrid As Boolean
    
    lRecriarGrid = False
    
    For i = 1 To flxCobertura.Rows - 1
        If flxCobertura.TextMatrix(i, 0) = Empty Then
            lRecriarGrid = True
            Exit For
        End If
    Next i
    
    If lRecriarGrid Then
        Call mpSelecionarFranquias(pColecaoPedCobert)
        
    Else
        With frmT46V117A.flxSelecionadas
            flxCobertura.Rows = .Rows
            For i = 1 To .Rows - 1
            
                j = mpGetFranquiaID(.TextMatrix(i, ColCodCobert))
                
                If j > 0 Then
                    flxCobertura.TextMatrix(j, 1) = .TextMatrix(i, ColDescFranquia)
                    If Trim(flxCobertura.TextMatrix(j, 1)) = Empty Then
                        flxCobertura.TextMatrix(j, 1) = .TextMatrix(i, ColNomCobert)
                    End If
                    flxCobertura.TextMatrix(j, 2) = Trim(.TextMatrix(i, ColTextoFranquia))
                    flxCobertura.TextMatrix(j, 3) = Trim(.TextMatrix(i, ColDesconto))
                End If
            Next
        End With
    End If
    
End Sub

Private Function mpGetObservacaoFranquia(ByVal CodCobert As String, ByVal ColecaoCoberturas As Collection) As String
    Dim iPedCobert As stcA46V704B
    Dim lObjstcPedCobertFranquia As stcA46V738B
    
    For Each iPedCobert In ColecaoCoberturas
        If iPedCobert.Cod_Cobert = CodCobert Then
        
            If Not ((iPedCobert.ObjPai.ObjPai.Num_Apol_An > 0 Or iPedCobert.ObjPai.ObjPai.objstcCtrlEmis.Num_Apol_Mar > 0) And iPedCobert.gMultiploFraq = 0 And iPedCobert.ObjPai.ObjPai.Cod_Produto >= 1080) Then
            If iPedCobert.ColecaoNovaFranquia.Count > 0 Then
                mpGetObservacaoFranquia = iPedCobert.ColecaoNovaFranquia(1).Texto_Observacoes
                End If
            Else
                Dim lrsFranquia As clsYasRecordSet
    
                Call gM46V999.mfFranqNovoLim(iPedCobert.Cod_Cobert, iPedCobert.Val_IS, frmT46V101A.objstcPedLoc.Cod_Clasf, frmT46V101A.objstcPedLoc.Cod_Compl_Clasf, _
                                             frmT46V101A.objstcPedLoc.Cod_Constr, Replace(frmT46V101A.txtRisCEP, "-", ""), frmT46V101A.cboTipCobert.ItemData(frmT46V101A.cboTipCobert.ListIndex), iPedCobert.Num_Meses_PI, lrsFranquia)
                If lrsFranquia.RecordCount > 0 Then
                    mpGetObservacaoFranquia = gM46V999.gfTratarNulo(lrsFranquia("Observacao_Franq"), tpoString)
                End If
            End If
            Exit Function
        End If
    Next
    mpGetObservacaoFranquia = Empty
End Function

Private Function mpGetDesconto(ByVal CodCobert As String, ByVal ColecaoCobert As Collection) As Integer
    Dim lobjstcCobert As stcA46V704B
    
    For Each lobjstcCobert In ColecaoCobert
        If lobjstcCobert.Cod_Cobert = CodCobert Then
            mpGetDesconto = lobjstcCobert.Per_Desc_Franq
            Exit Function
        End If
    Next
    mpGetDesconto = 0
End Function

Private Function mpGetNomeCobertura(ByVal CodCobert As String, ByVal ColecaoCobert As Collection) As String
    Dim lobjstcCobert As stcA46V704B
    
    For Each lobjstcCobert In ColecaoCobert
        If lobjstcCobert.Cod_Cobert = CodCobert Then
            mpGetNomeCobertura = lobjstcCobert.gNomeCobertura
            Exit Function
        End If
    Next
    mpGetNomeCobertura = Empty
End Function

Private Function mpGetNomeFranquia(ByVal CodCobert As String, ByVal ColecaoCobert As Collection) As String
    Dim lobjstcCobert As stcA46V704B
    
    For Each lobjstcCobert In ColecaoCobert
        If lobjstcCobert.Cod_Cobert = CodCobert Then
            mpGetNomeFranquia = lobjstcCobert.gNomeFranquia
            Exit Function
        End If
    Next
    
    mpGetNomeFranquia = Empty
End Function

Private Function mpExisteFranquia(ByVal pCodCobert As String) As Boolean
    Dim i As Integer
    
    For i = 1 To flxCobertura.Rows - 1
        If flxCobertura.TextMatrix(i, 0) = pCodCobert Then
            mpExisteFranquia = True
            Exit Function
        End If
    Next
    
    mpExisteFranquia = False
End Function

Public Sub mpExcluirFranquia(ByVal pCodCobert As String)
    Dim i As Integer
    
    For i = 1 To flxCobertura.Rows - 1
        If flxCobertura.TextMatrix(i, 0) = pCodCobert Then
            flxCobertura.RemoveItem (i)
            Exit Sub
        End If
    Next
End Sub

Public Sub mpIncluirFranquia(pCodCobert As String, pDescFranquia As String, pTextoFranquia As String, pDesconto As Double, _
                            pObservacaoFranquia As String)
    Dim i As Integer
    
    If Not mpExisteFranquia(pCodCobert) Then
        flxCobertura.Rows = flxCobertura.Rows + 1
        i = flxCobertura.Rows - 1
        
        'If pTextoFranquia <> Empty Then
            flxCobertura.TextMatrix(i, 0) = pCodCobert
            flxCobertura.TextMatrix(i, 1) = pDescFranquia
            flxCobertura.TextMatrix(i, 2) = Trim(pTextoFranquia)
            flxCobertura.TextMatrix(i, 3) = pDesconto
            flxCobertura.TextMatrix(i, 4) = pObservacaoFranquia
        'End If
    End If
End Sub

Public Function mpGetFranquiaID(ByVal pCodCobert As String) As Integer
    Dim i As Integer
    
    For i = 1 To flxCobertura.Rows - 1
        If flxCobertura.TextMatrix(i, 0) = pCodCobert Then
            mpGetFranquiaID = i
            Exit Function
        End If
    Next
    
    mpGetFranquiaID = 0
End Function
