VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT46V024A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selecione opção para nova cotação."
   ClientHeight    =   2430
   ClientLeft      =   11760
   ClientTop       =   450
   ClientWidth     =   2670
   Icon            =   "T46V024A.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   2670
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Opção para nova cotação"
      ForeColor       =   &H00800000&
      Height          =   1815
      Left            =   105
      TabIndex        =   2
      Top             =   120
      Width           =   2460
      Begin VB.OptionButton optOpcao 
         Caption         =   "Endosso"
         ForeColor       =   &H00800000&
         Height          =   400
         Index           =   2
         Left            =   135
         TabIndex        =   5
         Top             =   1215
         Width           =   1035
      End
      Begin VB.OptionButton optOpcao 
         Caption         =   "Renovação"
         ForeColor       =   &H00800000&
         Height          =   400
         Index           =   1
         Left            =   135
         TabIndex        =   4
         Top             =   727
         Width           =   1395
      End
      Begin VB.OptionButton optOpcao 
         Caption         =   "Seguro Novo"
         ForeColor       =   &H00800000&
         Height          =   400
         Index           =   0
         Left            =   135
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   1395
      End
   End
   Begin Threed.SSCommand cmdOk 
      Height          =   300
      Left            =   293
      TabIndex        =   0
      Top             =   2070
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Ok"
      ForeColor       =   8388608
   End
   Begin Threed.SSCommand cmdCancelar 
      Height          =   300
      Left            =   1373
      TabIndex        =   1
      Top             =   2070
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Cancelar"
      ForeColor       =   8388608
   End
End
Attribute VB_Name = "frmT46V024A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mNosNumero As String  'Nosso número informado pelos usuário de digitação.


Private Sub cmdCancelar_Click()
    gM46V999.gFormCancelado = True
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim lNossoNumero As String
    Dim lMensagem As String
    Dim lRegAlterados As Long

    Dim lobjCtrlEmis As clsA46V712A    'Objeto da classe clsA46V712A (TAB_CTRL_EMIS)
    Dim lobjstcCtrlEmis As stcA46V712B  'Objeto da classe de estrutura stcA46V712B (TAB_CTRL_EMIS)
    Dim lObjPed As clsA46V702A    'Objeto da classe clsA46V702A (TAB_PED).
    Dim lObjstcPed As stcA46V702B    'Objeto da classe de estrutura stcA46V702B (TAB_PED).
    Dim lobjstcPedLoc As stcA46V708B     'Classe (stc) de itens.
    Dim lobjPedLoc As clsA46V708A    'Objeto da classe clsA46V708A (TAB_PED_LOC).
    Dim lPropTabela As String

    Dim bdBasDados As ADODB.Connection    'Base de dados.

    'Call M46V999.gfObtSYASMULTINI(lPropTabela)
    
    If frmT46V012A.cboOpcao.ListIndex = 0 And optOpcao(0).Value = True Then
        Call gM46V999.gpGraLog(1, "Para este processo favor selecionar a opção (Demais).")
        frmT46V012A.txtPropostaSyas.Text = Empty
        frmT46V012A.txtDV.Text = Empty
        cmdCancelar_Click
        Exit Sub
    End If

    If gM46V999.gfAbrBasDados("P00MPED", gM46V999.gTipP00MPED, gM46V999.gEndP00MPED, bdBasDados, _
                              lMensagem) = False Then
        'mensagem
        Exit Sub
    End If

    '1. Posiciono o Nosso Numero
    lNossoNumero = ""
    If frmT46V012A.cboOpcao.ItemData(frmT46V012A.cboOpcao.ListIndex) <> 2 Then
        lNossoNumero = frmT46V012A.txtPropostaSyas.Text
    End If
    If optOpcao(0).Value = True Then    'Seguro Novo
        'Ficha 1000902
        If frmT46V012A.cboOpcao.ListIndex = 4 Then
            Call Shell(gM46V999.gEndExeP47V001 & " \" & Format(frmT46V012A.txtNumeroPI.Text, String(10, "0")) & "\" & Format(CDate(frmT46V012A.txtDataRecebimento), "yyyymmdd") & "\S", vbNormalFocus)
            Unload Me
            Exit Sub
        End If
        
        If frmT46V012A.cboOpcao.ItemData(frmT46V012A.cboOpcao.ListIndex) = 2 And Len(frmT46V012A.txtPropostaSyas.Text) > 12 Then
            Call gM46V999.gpGraLog(1, "FNP0044 - Na opção 'DEMAIS' o número do PROPOSTA-SYAS deve ser preenchido com o SISCOTA, caso não exista deixar sem preenchimento!")
            Exit Sub
        End If
        
        'Criar um seguro novo
        If Not gM46V001.gfSegNovo(lNossoNumero, "", 0, 0, 0, "") Then
            Exit Sub
        End If
        gM46V999.gFormCancelado = False
    ElseIf optOpcao(1).Value = True Then    'Renovação
        'Ficha 1000902
        If frmT46V012A.cboOpcao.ListIndex = 4 Then
            Call Shell(gM46V999.gEndExeP47V001 & " \" & Format(frmT46V012A.txtNumeroPI.Text, String(10, "0")) & "\" & Format(CDate(frmT46V012A.txtDataRecebimento), "yyyymmdd") & "\R", vbNormalFocus)
            Unload Me
            Exit Sub
        End If
        With frmT46V002A
            .Caption = "Cotação - Renovação"
            Call .mpInstanciarForm(lNossoNumero)
        End With
    Else  'Endosso
        'Ficha 1000902
        If frmT46V012A.cboOpcao.ListIndex = 4 Then
            Call Shell(gM46V999.gEndExeP47V001 & " \" & Format(frmT46V012A.txtNumeroPI.Text, String(10, "0")) & "\" & Format(CDate(frmT46V012A.txtDataRecebimento), "yyyymmdd") & "\E", vbNormalFocus)
            Unload Me
            Exit Sub
        End If
        With frmT46V178A
            Call .mpInstanciarForm(lNossoNumero)
        End With
    End If
    If Not gM46V999.gfPreenchido(lNossoNumero) Then
        Call gM46V999.gpGraLog(1, "FNP0045 - Nosso número não preenchido. Verifique")
        Exit Sub
    End If

    If gM46V999.gFormCancelado = False Then

        Set lObjPed = New clsA46V702A
        Set lObjPed.M46V999 = gM46V999

        Set lObjstcPed = New stcA46V702B
        Set lObjstcPed.M46V999 = gM46V999

        Set lobjCtrlEmis = New clsA46V712A
        Set lobjCtrlEmis.M46V999 = gM46V999

        Set lObjPed.ConexaoBD = bdBasDados

        lObjstcPed.NOSSO_NUMERO = lNossoNumero

        If Not lObjPed.mfCarregarColecao(lObjstcPed, "", "", True) Then
            Call gM46V999.gpGraLog(1, "FNP0043 - " & lObjPed.UltimoErro)
            Exit Sub
        End If

        Set lObjstcPed = lObjPed.Colecao.item(1)

        lObjstcPed.NUM_PI = frmT46V012A.txtNumeroPI
        lObjstcPed.Cod_Unicidade = Val(frmT46V012A.txtDigito.Text)
        lObjstcPed.Cod_Tip_Doc = e_TipDoc_Cotacao


        lObjstcPed.objstcCtrlEmis.NOSSO_NUMERO = lNossoNumero
        lObjstcPed.objstcCtrlEmis.NUM_PI = frmT46V012A.txtNumeroPI
        lObjstcPed.objstcCtrlEmis.Dat_Rec_Cia = gM46V999.gfForData(1, frmT46V012A.txtDataRecebimento)
        lObjstcPed.objstcCtrlEmis.Tip_Cotac = 0
        If frmT46V012A.cboOpcao.ItemData(frmT46V012A.cboOpcao.ListIndex) = 3 Then
            lObjstcPed.objstcCtrlEmis.Num_Ref_Corr = "MT" & Format(frmT46V012A.txtMTVVD.Text, "000000")
        End If
        'Ficha 1000449 - Incluir campo Referencia Marsh para digitação
        If frmT46V012A.cboOpcao.ItemData(frmT46V012A.cboOpcao.ListIndex) = 4 Then
            lObjstcPed.objstcCtrlEmis.Num_Ref_Corr = Format(frmT46V012A.txtMTVVD.Text, "000000")
        End If

        If frmT46V012A.cboOpcao.ItemData(frmT46V012A.cboOpcao.ListIndex) = 2 Then
            lObjstcPed.objstcCtrlEmis.Num_Cotac = frmT46V012A.txtPropostaSyas.Text
        End If

        'Vai verificar se existem dados na triagem se existir vai carregar tb.
        If Not gM46V001.mfCarregaTriagem(lObjstcPed, lMensagem, "") Then
            Call gM46V999.gpGraLog(1, "FNP0043 - " & lMensagem)
        End If

        If Not lobjCtrlEmis.mfAlterar(lObjstcPed.objstcCtrlEmis, "", "") Then
            Call gM46V999.gpGraLog(1, "FNP0042 - " & lobjCtrlEmis.UltimoErro)
            Exit Sub
        End If

        If Not lObjPed.mfAlterar(lObjstcPed, "", "", lObjstcPed.Dat_Liber, lObjstcPed.Cod_User_Liber) Then
            Call gM46V999.gpGraLog(1, "FNP0043 - " & lObjPed.UltimoErro)
            Exit Sub
        End If

        Set lobjPedLoc = New clsA46V708A
        Set lobjPedLoc.M46V999 = gM46V999

        For Each lobjstcPedLoc In lObjstcPed.ColecaoPedLoc
            If Not lobjPedLoc.mfAlterar(lobjstcPedLoc, "", "", False) Then
                Call gM46V999.gpGraLog(1, "FNP0043 - " & lobjPedLoc.UltimoErro)
                Exit Sub
            End If
        Next

        Set lObjPed = Nothing
        Set lobjPedLoc = Nothing
        Set lObjstcPed = Nothing
        Set lobjstcPedLoc = Nothing
        Set lobjCtrlEmis = Nothing
        Set lobjstcCtrlEmis = Nothing
        If App.EXEName = "P46V101" Then
            Unload Me
        Else
            If gfShell("P46V100") = True Then
                Unload Me
            End If
        End If
    Else
        Unload Me
    End If
End Sub

Private Sub Form_Activate()
    frmT46V001A.panForm = Mid$(Me.Name, 4)
End Sub

Private Sub Form_Load()
    Call gM46V999.gpCenForm(Me)
    gM46V999.gFormCancelado = False
End Sub

Private Sub optOpcao_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdOk_Click
    End If
End Sub

