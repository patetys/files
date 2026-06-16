VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmT46V103A 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selecione a região"
   ClientHeight    =   3690
   ClientLeft      =   2745
   ClientTop       =   2175
   ClientWidth     =   7530
   ControlBox      =   0   'False
   Icon            =   "T46V103A.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid flxRegiao 
      Height          =   3135
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Região."
      Top             =   120
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   5530
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCommand cmdOk 
      Height          =   300
      Left            =   3480
      TabIndex        =   0
      ToolTipText     =   "Volta ao item com a região selecionada."
      Top             =   3360
      Width           =   1020
      _Version        =   65536
      _ExtentX        =   1799
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Ok"
      ForeColor       =   8388608
   End
End
Attribute VB_Name = "frmT46V103A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lSortCriteria As Long    'Utilizado para ordenação do grid
Dim lLastColSort As Long    'Utilizado para ordenação do grid
Private Sub cmdOK_Click()
    If flxRegiao.RowSel <> 0 Then
        frmT46V101A.txtRisRegiao.Tag = Format(flxRegiao.TextMatrix(flxRegiao.RowSel, 0), "00000") & Format(flxRegiao.TextMatrix(flxRegiao.RowSel, 2), "00000")
        gM46V111.gCod_Regiao = Val(Right(frmT46V101A.txtRisRegiao.Tag, 4))
        frmT46V101A.txtRisRegiao.Text = flxRegiao.TextMatrix(flxRegiao.RowSel, 1)
    End If
    Unload Me
End Sub

Private Sub Form_Activate()
'1. Posiciona controles.
    frmT46V100A.panForm.Caption = Mid$(Me.Name, 4)
End Sub

Private Sub Form_Load()
    Dim lobjstc As stcA46V101B      'Objeto da classe de estrutura stcA46V101A.
    Dim lbdTAB_Regiao As ADODB.Connection       'Arquivo P00Tarifa.
    Dim lMensagem As String             'Mensagem.
    Dim lrsTab_Regiao As clsYasRecordSet        'Registro TAB_REGIAO.
    Dim lLinha As Integer         'Linha.
    Dim lSelect As String
    If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTAB_Regiao, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If

    lSelect = " SELECT TAB_REGIAO.DSC_REGIAO, TAB_REGIAO.COD_REGIAO,TAB_MULTI_REGIAO.COD_PRODUTO "
    lSelect = lSelect & " FROM TAB_MULTI_REGIAO INNER JOIN TAB_REGIAO ON "
    lSelect = lSelect & " TAB_MULTI_REGIAO.COD_REGIAO = TAB_REGIAO.COD_REGIAO"
    lSelect = lSelect & " Where cod_Ramo = " & gM46V999.gCod_Ramo
    lSelect = lSelect & " and DAT_TER_VIG = 99999999"
    lSelect = lSelect & " order by TAB_REGIAO.cod_regiao"

    If gM46V999.gfObtRegistro(lbdTAB_Regiao, lSelect, lrsTab_Regiao, lMensagem) = False Then
        Call gM46V999.gpFecha2(lbdTAB_Regiao)
        Call gM46V999.gpGraLog(2, lMensagem)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    flxRegiao.Rows = 1
    flxRegiao.Cols = 2
    flxRegiao.Cols = 3
    flxRegiao.TextMatrix(0, 0) = "Código"
    flxRegiao.TextMatrix(0, 1) = "Descrição"
    flxRegiao.TextMatrix(0, 2) = "cod_produto"
    flxRegiao.ColWidth(0) = 650
    flxRegiao.ColWidth(1) = 20000
    flxRegiao.ColWidth(2) = 0
    Do While Not lrsTab_Regiao.EOF
        flxRegiao.Rows = flxRegiao.Rows + 1
        flxRegiao.TextMatrix(flxRegiao.Rows - 1, 0) = Format(lrsTab_Regiao("COD_REGIAO"), "000")
        flxRegiao.TextMatrix(flxRegiao.Rows - 1, 1) = lrsTab_Regiao("DSC_REGIAO")
        flxRegiao.TextMatrix(flxRegiao.Rows - 1, 2) = lrsTab_Regiao("COD_PRODUTO")
        lrsTab_Regiao.MoveNext
    Loop
End Sub

Private Sub flxRegiao_DblClick()
    If flxRegiao.MouseRow = 0 Then
        If lLastColSort <> flxRegiao.Col Then
            lSortCriteria = 1
        Else
            lSortCriteria = IIf(lSortCriteria = 1, 2, 1)
        End If

        flxRegiao.ColSel = flxRegiao.Col
        lLastColSort = flxRegiao.Col
        flxRegiao.Sort = lSortCriteria
        If flxRegiao.Rows = 1 Then
            Call cmdOK_Click
        End If
    Else
        Call cmdOK_Click
    End If
End Sub
