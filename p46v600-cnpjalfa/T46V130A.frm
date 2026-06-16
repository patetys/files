VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmT46V130A 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1680
   ClientLeft      =   1770
   ClientTop       =   3555
   ClientWidth     =   6165
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1680
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid flxTabela 
      Height          =   1545
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6230
      _ExtentX        =   10980
      _ExtentY        =   2725
      _Version        =   393216
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "frmT46V130A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mValSelecionado As Integer

Private Sub flxTabela_DblClick()
    If flxTabela.RowSel = 0 Or flxTabela.ColSel = 0 Then
        mValSelecionado = 0
    Else
        If Trim(flxTabela.TextMatrix(flxTabela.RowSel, flxTabela.ColSel)) = "-" Then
            mValSelecionado = 0
        Else
            mValSelecionado = Trim(Replace(flxTabela.TextMatrix(flxTabela.RowSel, flxTabela.ColSel), "%", ""))
        End If
    End If
    Unload Me
End Sub

Private Sub Form_Load()
'Call gM46V999.gpCenForm(Me)
    With flxTabela
        .Cols = 6
        .Rows = 6
        .FixedCols = 1
        .FixedRows = 1
        .TextMatrix(0, 0) = "  Sin./Premio  "
        .TextMatrix(0, 1) = "   12 meses    "
        .TextMatrix(0, 2) = "   24 meses    "
        .TextMatrix(0, 3) = "   36 meses    "
        .TextMatrix(0, 4) = "   48 meses    "
        .TextMatrix(0, 5) = "   60 meses    "
        .TextMatrix(1, 0) = "    S/P=0%    "
        .TextMatrix(2, 0) = " 0%< S/P <=10%"
        .TextMatrix(3, 0) = "10%< S/P <=20%"
        .TextMatrix(4, 0) = "20%< S/P <=30%"
        .TextMatrix(5, 0) = "30%< S/P <=40%"
        .ColWidth(0) = 0    '1200539 - Tarifa 1038 e 1039 Junho 2012 - Item 5
        If gM46V999.gCodProduto >= 932 And frmT46V112A.optTipEmissao(1).Value = True And _
           gM46V999.gCod_Ramo = e_CodRamo_Empresarial Then
            .TextMatrix(1, 1) = "      10%      "
            .TextMatrix(2, 1) = "       5%      "
            .TextMatrix(3, 1) = "       5%      "
            .TextMatrix(4, 1) = "       5%      "
            .TextMatrix(5, 1) = "       5%      "
            .TextMatrix(1, 2) = "      15%      "
            .TextMatrix(2, 2) = "      10%      "
            .TextMatrix(3, 2) = "       5%      "
            .TextMatrix(4, 2) = "       5%      "
            .TextMatrix(5, 2) = "       5%      "
            .TextMatrix(1, 3) = "      20%      "
            .TextMatrix(2, 3) = "      15%      "
            .TextMatrix(3, 3) = "      10%      "
            .TextMatrix(4, 3) = "       5%      "
            .TextMatrix(5, 3) = "       5%      "
            .TextMatrix(1, 4) = "      25%      "
            .TextMatrix(2, 4) = "      20%      "
            .TextMatrix(3, 4) = "      15%      "
            .TextMatrix(4, 4) = "      10%      "
            .TextMatrix(5, 4) = "       5%      "
            .TextMatrix(1, 5) = "      30%      "
            .TextMatrix(2, 5) = "      25%      "
            .TextMatrix(3, 5) = "      20%      "
            .TextMatrix(4, 5) = "      15%      "
            .TextMatrix(5, 5) = "      10%      "
        ElseIf gM46V999.gCodProduto >= 946 And frmT46V112A.optTipEmissao(1).Value = True And _
               gM46V999.gCod_Ramo = e_CodRamo_Residencial Then
            .TextMatrix(1, 1) = "      10%      "
            .TextMatrix(2, 1) = "       5%      "
            .TextMatrix(3, 1) = "       5%      "
            .TextMatrix(4, 1) = "       5%      "
            .TextMatrix(5, 1) = "       5%      "
            .TextMatrix(1, 2) = "      15%      "
            .TextMatrix(2, 2) = "      10%      "
            .TextMatrix(3, 2) = "       5%      "
            .TextMatrix(4, 2) = "       5%      "
            .TextMatrix(5, 2) = "       5%      "
            .TextMatrix(1, 3) = "      20%      "
            .TextMatrix(2, 3) = "      15%      "
            .TextMatrix(3, 3) = "      10%      "
            .TextMatrix(4, 3) = "       5%      "
            .TextMatrix(5, 3) = "       5%      "
            .TextMatrix(1, 4) = "      25%      "
            .TextMatrix(2, 4) = "      20%      "
            .TextMatrix(3, 4) = "      15%      "
            .TextMatrix(4, 4) = "      10%      "
            .TextMatrix(5, 4) = "       5%      "
            .TextMatrix(1, 5) = "      30%      "
            .TextMatrix(2, 5) = "      25%      "
            .TextMatrix(3, 5) = "      20%      "
            .TextMatrix(4, 5) = "      15%      "
            .TextMatrix(5, 5) = "      10%      "
        Else
            .TextMatrix(1, 1) = "       5%      "
            .TextMatrix(2, 1) = "       -       "
            .TextMatrix(3, 1) = "       -       "
            .TextMatrix(4, 1) = "       -       "
            .TextMatrix(5, 1) = "       -       "
            .TextMatrix(1, 2) = "      10%      "
            .TextMatrix(2, 2) = "       5%      "
            .TextMatrix(3, 2) = "       -       "
            .TextMatrix(4, 2) = "       -       "
            .TextMatrix(5, 2) = "       -       "
            .TextMatrix(1, 3) = "      15%      "
            .TextMatrix(2, 3) = "      10%      "
            .TextMatrix(3, 3) = "       5%      "
            .TextMatrix(4, 3) = "       -       "
            .TextMatrix(5, 3) = "       -       "
            .TextMatrix(1, 4) = "      20%      "
            .TextMatrix(2, 4) = "      15%      "
            .TextMatrix(3, 4) = "      10%      "
            .TextMatrix(4, 4) = "       5%      "
            .TextMatrix(5, 4) = "      30%      "
            .TextMatrix(1, 5) = "      25%      "
            .TextMatrix(2, 5) = "      20%      "
            .TextMatrix(3, 5) = "      15%      "
            .TextMatrix(4, 5) = "      10%      "
            .TextMatrix(5, 5) = "       5%      "
        End If
        
        '8000046257, Equiparação das tabelas de bônus
        If gM46V999.gCodProduto >= 1166 And _
               gM46V999.gCod_Ramo = e_CodRamo_Residencial Then
            .TextMatrix(1, 1) = "       5%      "
            .TextMatrix(2, 1) = "       -       "
            .TextMatrix(3, 1) = "       -       "
            .TextMatrix(4, 1) = "       -       "
            .TextMatrix(5, 1) = "       -       "
            .TextMatrix(1, 2) = "      10%      "
            .TextMatrix(2, 2) = "       5%      "
            .TextMatrix(3, 2) = "       -       "
            .TextMatrix(4, 2) = "       -       "
            .TextMatrix(5, 2) = "       -       "
            .TextMatrix(1, 3) = "      15%      "
            .TextMatrix(2, 3) = "      10%      "
            .TextMatrix(3, 3) = "       5%      "
            .TextMatrix(4, 3) = "       -       "
            .TextMatrix(5, 3) = "       -       "
            .TextMatrix(1, 4) = "      20%      "
            .TextMatrix(2, 4) = "      15%      "
            .TextMatrix(3, 4) = "      10%      "
            .TextMatrix(4, 4) = "       5%      "
            .TextMatrix(5, 4) = "      30%      "
            .TextMatrix(1, 5) = "      25%      "
            .TextMatrix(2, 5) = "      20%      "
            .TextMatrix(3, 5) = "      15%      "
            .TextMatrix(4, 5) = "      10%      "
            .TextMatrix(5, 5) = "       5%      "
        End If
        
    End With
End Sub
Public Sub mpInstanciarForm(pValor As Integer, pOwner As Form)
    Me.Show vbModal, pOwner
    pValor = mValSelecionado
End Sub
