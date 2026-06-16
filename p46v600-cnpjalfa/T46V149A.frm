VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmT46V149A 
   Caption         =   "Alteração de prêmio e prêmio Anual"
   ClientHeight    =   4770
   ClientLeft      =   2745
   ClientTop       =   1470
   ClientWidth     =   7395
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4770
   ScaleWidth      =   7395
   Begin MSFlexGridLib.MSFlexGrid flxSelecionadas 
      Height          =   4320
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   7620
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
      GridLines       =   2
      AllowUserResizing=   1
   End
   Begin Threed.SSCommand cmdOk 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   300
      Left            =   3375
      TabIndex        =   1
      ToolTipText     =   "Volta à janela de item."
      Top             =   4455
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1764
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Ok"
      ForeColor       =   8388608
   End
End
Attribute VB_Name = "frmT46V149A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mfLimpaGrid()
    With flxSelecionadas
        .Rows = 1
        .Cols = 9
        .TextMatrix(0, 0) = "Cód"
        .TextMatrix(0, 1) = "Cobertura"
        .TextMatrix(0, 2) = "Prêmio"
        .TextMatrix(0, 3) = "Prêmio Anual"
        .TextMatrix(0, 4) = "Prêmio Ress"
        .TextMatrix(0, 5) = "Prêmio Ress Anual "
        .TextMatrix(0, 6) = "Prêmio Tarif"
        .TextMatrix(0, 7) = "Prêmio Tarif Anual"
        .TextMatrix(0, 8) = "Flag"
        .Row = 0
        .ColWidth(0) = 375
        .ColWidth(1) = 3500
        .ColWidth(2) = 1500
        .ColWidth(3) = 1500
        .ColWidth(4) = 0
        .ColWidth(5) = 0
        .ColWidth(6) = 0
        .ColWidth(7) = 0
        .ColWidth(8) = 0
    End With

End Sub

Private Sub cmdOk_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
    Call gM46V999.gpCenForm(Me)
    frmT46V100A.panForm.Caption = Mid$(Me.Name, 4)
    Call mfLimpaGrid
End Sub

Public Sub mpCarregaCoberturas(pColecaoPedCobert As Collection)
    Dim lobjstcCobert As stcA46V704B        'Objeto da classe de estrutura stcA46V704B (TAB_PED_COBERT).
    Call mfLimpaGrid
    flxSelecionadas.Rows = 1
    For Each lobjstcCobert In pColecaoPedCobert
        flxSelecionadas.Rows = flxSelecionadas.Rows + 1
        flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, 0) = lobjstcCobert.Cod_Cobert
        flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, 1) = lobjstcCobert.gNomeCobertura
        flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, 2) = Format(lobjstcCobert.VAL_PR, "standard")
        flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, 3) = Format(lobjstcCobert.Val_Pr_Anu, "standard")
        flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, 4) = Format(lobjstcCobert.Val_Pr_Resg, "standard")
        flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, 5) = Format(lobjstcCobert.Val_Pr_Resg_Anu, "standard")
        flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, 6) = Format(lobjstcCobert.Val_Pr_Tarif, "standard")
        flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, 7) = Format(lobjstcCobert.Val_Pr_Tarif_Anu, "standard")
    Next lobjstcCobert

End Sub



