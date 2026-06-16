VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmT46V117A 
   Caption         =   "Coberturas"
   ClientHeight    =   6480
   ClientLeft      =   1605
   ClientTop       =   1065
   ClientWidth     =   9960
   ControlBox      =   0   'False
   Icon            =   "T46V117A.frx":0000
   LockControls    =   -1  'True
   ScaleHeight     =   6480
   ScaleWidth      =   9960
   Begin VB.ComboBox CBO_ordenaCob 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "T46V117A.frx":000C
      Left            =   8460
      List            =   "T46V117A.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2865
      Width           =   1485
   End
   Begin VB.ListBox LstFranquias 
      Height          =   1230
      Left            =   8640
      TabIndex        =   8
      Top             =   3960
      Visible         =   0   'False
      Width           =   1170
   End
   Begin MSFlexGridLib.MSFlexGrid flxSelecionadas 
      Height          =   2145
      Left            =   45
      TabIndex        =   7
      Top             =   3510
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   3784
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
   Begin Threed.SSCommand cmdAdicionar 
      Height          =   300
      Left            =   3300
      TabIndex        =   2
      ToolTipText     =   "Contrata a(s) cobertura(s) selecionada(s)."
      Top             =   2850
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Adicionar"
      ForeColor       =   8388608
   End
   Begin Threed.SSCommand cmdRetirar 
      Height          =   300
      Left            =   4515
      TabIndex        =   3
      ToolTipText     =   "Retira a cobertura cobtratada."
      Top             =   2850
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Retirar"
      ForeColor       =   8388608
   End
   Begin MSFlexGridLib.MSFlexGrid flxCobertura 
      Height          =   2340
      Left            =   60
      TabIndex        =   1
      Top             =   405
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
   Begin VB.Frame fraCobertura 
      Height          =   810
      Left            =   0
      TabIndex        =   12
      Top             =   5670
      Visible         =   0   'False
      Width           =   4455
      Begin VB.TextBox txtDescExp 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1740
         MaxLength       =   5
         TabIndex        =   16
         Text            =   "0,00"
         ToolTipText     =   "Desconto de experiência."
         Top             =   465
         Visible         =   0   'False
         Width           =   1035
      End
      Begin Threed.SSCommand cmdMover 
         Height          =   255
         Left            =   2835
         TabIndex        =   15
         ToolTipText     =   "Move o desconto comercial para todas as coberturas contratadas."
         Top             =   165
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Mover para todas"
         ForeColor       =   8388608
      End
      Begin VB.TextBox txtDescComercial 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1740
         MaxLength       =   5
         TabIndex        =   13
         Text            =   "0,00"
         ToolTipText     =   "Desconto comercial."
         Top             =   150
         Width           =   1035
      End
      Begin Threed.SSCommand cmdBonus 
         Height          =   240
         Left            =   2835
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Exibir bem."
         Top             =   495
         Visible         =   0   'False
         Width           =   240
         _Version        =   65536
         _ExtentX        =   423
         _ExtentY        =   423
         _StockProps     =   78
         ForeColor       =   12582912
         Picture         =   "T46V117A.frx":002D
      End
      Begin VB.Label lblMulti 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Desconto experiência: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   45
         TabIndex        =   17
         Top             =   465
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.Label lblMulti 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Desconto comercial: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   20
         Left            =   45
         TabIndex        =   14
         Top             =   150
         Width           =   1680
      End
   End
   Begin Threed.SSCommand cmdFranquia 
      Height          =   345
      Left            =   7770
      TabIndex        =   11
      ToolTipText     =   "Calcula as coberturas e volta ao item"
      Top             =   5790
      Width           =   2115
      _Version        =   65536
      _ExtentX        =   3731
      _ExtentY        =   609
      _StockProps     =   78
      Caption         =   "Busca Franquia Mínima"
      ForeColor       =   8388608
   End
   Begin MSComctlLib.TabStrip TabStrip2 
      Height          =   2610
      Left            =   0
      TabIndex        =   6
      Top             =   3120
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   4604
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Coberturas contratadas"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCommand cmdOK 
      Height          =   345
      Left            =   4418
      TabIndex        =   9
      ToolTipText     =   "Calcula as coberturas e volta ao item"
      Top             =   5775
      Width           =   1125
      _Version        =   65536
      _ExtentX        =   1984
      _ExtentY        =   609
      _StockProps     =   78
      Caption         =   "&Ok"
      ForeColor       =   8388608
   End
   Begin Threed.SSCommand cmdTexto 
      Height          =   345
      Left            =   5610
      TabIndex        =   10
      ToolTipText     =   "Calcula as coberturas e volta ao item"
      Top             =   5790
      Width           =   2115
      _Version        =   65536
      _ExtentX        =   3731
      _ExtentY        =   609
      _StockProps     =   78
      Caption         =   "Ver Texto Franquia"
      ForeColor       =   8388608
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   2790
      Left            =   45
      TabIndex        =   0
      Top             =   15
      Width           =   9900
      _ExtentX        =   17463
      _ExtentY        =   4921
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Coberturas"
            Object.ToolTipText     =   "Clique duas vezes sobre a cobertura desejada ou selecione as coberturas e clique em Adicionar."
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label LBL_ordenaCob 
      Caption         =   "Ordenar por:"
      Height          =   300
      Left            =   7530
      TabIndex        =   4
      Top             =   2910
      Width           =   960
   End
End
Attribute VB_Name = "frmT46V117A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ColecaoCobert As Collection    'Coleção.
Dim mLinhaFocus As Long    'Linha que recebeu o foco.
Dim mColunaFocus As Long    'Coluna que recebeu o foco.
Dim mLinhaInicial As Long    'Linha inicial (seleção).
Dim mLinhaFinal As Long    'Linha final (seleção).
Dim lSortCriteria As Long    'Utilizado para ordenação do grid
Dim lLastColSort As Long    'Utilizado para ordenação do grid
Dim lSortCriteriaSel As Long  'Utilizado para ordenação do grid
Dim lLastColSortSel As Long    'Utilizado para ordenação do grid
Dim lPlanoSimples As Boolean
Dim lTip_Atividade As Integer    'Utilizado para selecao de Franquia
Dim PressUpDown As Boolean    'Verificar se pressionou Teclas Up ou Down
Dim RowIsAlterada As Integer    'Caso a IS seja alterado vai guardar a linha para fazer acerto futuro
Dim RowAlterouFranquia As Integer
Dim mCobertSelec As String
Dim TrocandoCor As Boolean

Private mColecaoAdicional As Collection
Private mUltPlanoA As Integer
Private mUltProdutoA As Integer

Private mColecaoBasica As Collection
Private mUltValorLMGB As Double
Private mUltPlanoB As Integer
Private mUltProdutoB As Integer
Private blnRECALCULA_MULTIPLO As Boolean

Public gCepAlterado As Boolean
'CR23794 - Syas Kit - Erro PI
Public gPiDefault As Boolean

Sub DesableFranquia()
    LstFranquias.Visible = False
End Sub

Function mfBuscaCoefFranquiasBasica(ByVal ValorLMG As Double, ByVal pValorFranquia As Double) As Double
    Dim lbdTbm_Franquia_Basica As ADODB.Connection   'Arquivo: P00Tarifa.
    Dim lMensagem As String    'Mensagem.
    Dim lrsTbm_Franquia_Basica As clsYasRecordSet    'Registro : Tbm_Franquia_Basica.
    Dim Aux As String
    Dim lobjstcAuxiliar As stcAuxiliar
    Dim lFatorAjuste As Double
    Dim lSelect As String
    If Me.Tag = "Desabilitado" Then Exit Function

    If frmT46V101A.cboPlano.ListIndex = -1 Then
        Exit Function
    End If

    If mColecaoBasica Is Nothing Then
        Set mColecaoBasica = New Collection
        mUltPlanoB = 0
        mUltProdutoB = 0
        mUltValorLMGB = 0
    End If

    If mUltPlanoB <> frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex) Or _
       mUltProdutoB <> gM46V999.gCodProduto Or _
       mUltValorLMGB <> ValorLMG Then
        If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTbm_Franquia_Basica, lMensagem) = False Then
            Call gM46V999.gpGraLog(2, lMensagem)
            Exit Function
        End If

        lSelect = " Select distinct Val_Franq, Fator_Ajuste from Tbm_Franquia_Basica "
        lSelect = lSelect & " Where cod_Ramo = " & gM46V999.gCod_Ramo
        lSelect = lSelect & " and Cod_Modalidade = " & frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex)
        lSelect = lSelect & " and Val_LMG_Min <= " & gM46V999.gfForValor(1, ValorLMG)
        If gM46V999.gCodProduto >= 942 Then
            lSelect = lSelect & " and Val_LMG_Max > " & gM46V999.gfForValor(1, ValorLMG)
        Else
            lSelect = lSelect & " and Val_LMG_Max >= " & gM46V999.gfForValor(1, ValorLMG)
        End If
        lSelect = lSelect & " and Cod_Tarif = " & gM46V999.gCodProduto
        lSelect = lSelect & " and Cod_Sequencia <= " & gM46V999.gCodSequencia

        If gM46V999.gfObtRegistro(lbdTbm_Franquia_Basica, lSelect, lrsTbm_Franquia_Basica, lMensagem) = False Then
            Call gM46V999.gpGraLog(2, lMensagem)
            Call gM46V999.gpFecha2(lbdTbm_Franquia_Basica)
            Screen.MousePointer = vbDefault
            Exit Function
        End If

        Set mColecaoBasica = New Collection
        If Not lrsTbm_Franquia_Basica.EOF Then
            Do While Not lrsTbm_Franquia_Basica.EOF
                Set lobjstcAuxiliar = New stcAuxiliar
                lobjstcAuxiliar.Val_Fraquia = lrsTbm_Franquia_Basica("Val_Franq")
                lobjstcAuxiliar.Fator_Ajuste = lrsTbm_Franquia_Basica("Fator_Ajuste")
                mColecaoBasica.Add lobjstcAuxiliar
                lrsTbm_Franquia_Basica.MoveNext
            Loop
        End If

        mUltPlanoB = frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex)
        mUltProdutoB = gM46V999.gCodProduto
        mUltValorLMGB = ValorLMG

        Call gM46V999.gpFecha2(lbdTbm_Franquia_Basica)
    End If

    lFatorAjuste = 99
    If pValorFranquia > 0 Then
        For Each lobjstcAuxiliar In mColecaoBasica
            If lobjstcAuxiliar.Val_Fraquia <= pValorFranquia Then
                If lFatorAjuste >= lobjstcAuxiliar.Fator_Ajuste Then
                    lFatorAjuste = lobjstcAuxiliar.Fator_Ajuste
                End If
            End If
        Next lobjstcAuxiliar
    End If
    If lFatorAjuste <> 99 Then
        mfBuscaCoefFranquiasBasica = lFatorAjuste
    End If
End Function

Sub mpMostraFranquiasBasica(ByVal ValorLMG As Double)
    Dim lbdTbm_Franquia_Basica As ADODB.Connection   'Arquivo: P00Tarifa.
    Dim lMensagem As String    'Mensagem.
    Dim lrsTbm_Franquia_Basica As clsYasRecordSet    'Registro : Tbm_Franquia_Basica.
    Dim Aux As String
    Dim lobjstcAuxiliar As stcAuxiliar
    Dim lSelect As String
    If Me.Tag = "Desabilitado" Then Exit Sub


    If frmT46V101A.cboPlano.ListIndex = -1 Then
        Exit Sub
    End If

    If gM46V999.gNovaPropostaLMI Then
        Exit Sub  'Para LMI não permitira mostrar a lista
    End If


    If mColecaoBasica Is Nothing Then
        Set mColecaoBasica = New Collection
        mUltPlanoB = 0
        mUltProdutoB = 0
        mUltValorLMGB = 0
    End If

    If mUltPlanoB <> frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex) Or _
       mUltProdutoB <> gM46V999.gCodProduto Or _
       mUltValorLMGB <> ValorLMG Then
        If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTbm_Franquia_Basica, lMensagem) = False Then
            Call gM46V999.gpGraLog(2, lMensagem)
            Exit Sub
        End If


        lSelect = " Select distinct Val_Franq, Fator_Ajuste from Tbm_Franquia_Basica "
        lSelect = lSelect & " Where cod_Ramo = " & gM46V999.gCod_Ramo
        lSelect = lSelect & " and Cod_Modalidade = " & frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex)
        lSelect = lSelect & " and Val_LMG_Min <= " & gM46V999.gfForValor(1, ValorLMG)
        If gM46V999.gCodProduto >= 942 Then
            lSelect = lSelect & " and Val_LMG_Max > " & gM46V999.gfForValor(1, ValorLMG)
        Else
            lSelect = lSelect & " and Val_LMG_Max >= " & gM46V999.gfForValor(1, ValorLMG)
        End If
        lSelect = lSelect & " and Cod_Tarif = " & gM46V999.gCodProduto
        lSelect = lSelect & " and Cod_Sequencia <= " & gM46V999.gCodSequencia

        If gM46V999.gfObtRegistro(lbdTbm_Franquia_Basica, lSelect, lrsTbm_Franquia_Basica, lMensagem) = False Then
            Call gM46V999.gpGraLog(2, lMensagem)
            Call gM46V999.gpFecha2(lbdTbm_Franquia_Basica)
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        Set mColecaoBasica = New Collection
        If Not lrsTbm_Franquia_Basica.EOF Then
            Do While Not lrsTbm_Franquia_Basica.EOF
                Set lobjstcAuxiliar = New stcAuxiliar
                lobjstcAuxiliar.Val_Fraquia = lrsTbm_Franquia_Basica("Val_Franq")
                lobjstcAuxiliar.Fator_Ajuste = lrsTbm_Franquia_Basica("Fator_Ajuste")
                mColecaoBasica.Add lobjstcAuxiliar
                lrsTbm_Franquia_Basica.MoveNext
            Loop
        Else
            Call mpMostraFranquiasAdicional(ValorLMG)
        End If

        mUltPlanoB = frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex)
        mUltProdutoB = gM46V999.gCodProduto
        mUltValorLMGB = ValorLMG

        Call gM46V999.gpFecha2(lbdTbm_Franquia_Basica)
    End If


    For Each lobjstcAuxiliar In mColecaoBasica
        Aux = Format(lobjstcAuxiliar.Val_Fraquia, "standard")
        If gM46V999.CDblx(flxSelecionadas.TextMatrix(flxSelecionadas.Row, ColValFranMin)) < lobjstcAuxiliar.Val_Fraquia Then
            Aux = Space(12 - Len(Aux)) & Aux
            LstFranquias.AddItem Aux & "            - 0" '& lobjstcAuxiliar.Fator_Ajuste
        End If
    Next lobjstcAuxiliar
End Sub


Function mfBuscaCoefFranquiasAdicional(ByVal pValorLMG As Double, ByVal pValorFranquia As Double) As Double
    Dim lbdTbm_Franquia_Adicional As ADODB.Connection   'Arquivo: P00Tarifa.
    Dim lMensagem As String    'Mensagem.
    Dim lrsTbm_Franquia_Adicional As clsYasRecordSet    'Registro : Tbm_Franquia_Adicional.
    Dim Aux As String
    Dim lCodSequencia As Integer
    Dim lFatorAjuste As Double
    Dim lSelect As String

    If Me.Tag = "Desabilitado" Then Exit Function


    If mColecaoAdicional Is Nothing Then
        Set mColecaoAdicional = New Collection
        mUltPlanoA = 0
        mUltProdutoA = 0
    End If
    If frmT46V101A.cboPlano.ListIndex = -1 Then
        Exit Function
    End If

    If mUltProdutoA <> gM46V999.gCodProduto Or mUltPlanoA <> frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex) Then
        If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTbm_Franquia_Adicional, lMensagem) = False Then
            Call gM46V999.gpGraLog(2, lMensagem)
            Exit Function
        End If

        lCodSequencia = gM46V999.gCodSequencia

NovaSequencia:

        lSelect = " Select distinct Perc_Franquia, Fator_Ajuste from Tbm_Franquia_Adicional "
        lSelect = lSelect & " Where Cod_Ramo = " & gM46V999.gCod_Ramo
        lSelect = lSelect & " and Cod_Modalidade = " & frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex)
        lSelect = lSelect & " and Cod_Tarif     = " & gM46V999.gCodProduto
        lSelect = lSelect & " and Cod_Sequencia <= " & lCodSequencia
        lSelect = lSelect & " order by  Perc_Franquia "

        If gM46V999.gfObtRegistro(lbdTbm_Franquia_Adicional, lSelect, lrsTbm_Franquia_Adicional, lMensagem) = False Then
            Call gM46V999.gpFecha2(lbdTbm_Franquia_Adicional)
            Call gM46V999.gpGraLog(2, lMensagem)
            Screen.MousePointer = vbDefault
            Exit Function
        End If

        If lrsTbm_Franquia_Adicional.EOF Then
            If lCodSequencia > 0 Then
                lCodSequencia = lCodSequencia - 1
                GoTo NovaSequencia
            End If
        End If

        Set mColecaoAdicional = New Collection
        Dim lobjstcAux As stcAuxiliar
        Do While Not lrsTbm_Franquia_Adicional.EOF
            Set lobjstcAux = New stcAuxiliar
            lobjstcAux.Perc_Franquia = lrsTbm_Franquia_Adicional("Perc_Franquia")
            lobjstcAux.Fator_Ajuste = lrsTbm_Franquia_Adicional("Fator_Ajuste")
            mColecaoAdicional.Add lobjstcAux
            lrsTbm_Franquia_Adicional.MoveNext
        Loop

        mUltProdutoA = gM46V999.gCodProduto
        mUltPlanoA = frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex)

        Call gM46V999.gpFecha2(lbdTbm_Franquia_Adicional)
    End If
    lFatorAjuste = 99
    If pValorFranquia > 0 Then
        For Each lobjstcAux In mColecaoAdicional
            If gM46V999.gCodProduto >= 1080 Then
                If lFatorAjuste >= lobjstcAux.Fator_Ajuste Then
                    lFatorAjuste = lobjstcAux.Fator_Ajuste
                End If
            Else
                Aux = (lobjstcAux.Perc_Franquia) / 100 * pValorLMG
                If Aux <= pValorFranquia Then
                    If lFatorAjuste >= lobjstcAux.Fator_Ajuste Then
                        lFatorAjuste = lobjstcAux.Fator_Ajuste
                    End If
                End If
            End If
        Next lobjstcAux
    End If
    If lFatorAjuste <> 99 Then
        mfBuscaCoefFranquiasAdicional = lFatorAjuste
    End If

End Function


Sub mpMostraFranquiasAdicional(ByVal ValorLMG As Double)
    Dim lbdTbm_Franquia_Adicional As ADODB.Connection   'Arquivo: P00Tarifa.
    Dim lMensagem As String    'Mensagem.
    Dim lrsTbm_Franquia_Adicional As clsYasRecordSet    'Registro : Tbm_Franquia_Adicional.
    Dim Aux As String
    Dim lCodSequencia As Integer
    Dim lSelect As String

    If Me.Tag = "Desabilitado" Then Exit Sub

    If gM46V999.gNovaPropostaLMI Then
        Exit Sub  'Para LMI não permitira mostrar a lista
    End If

    If mColecaoAdicional Is Nothing Then
        Set mColecaoAdicional = New Collection
        mUltPlanoA = 0
        mUltProdutoA = 0
    End If
    If frmT46V101A.cboPlano.ListIndex = -1 Then
        Exit Sub
    End If

    If mUltProdutoA <> gM46V999.gCodProduto Or mUltPlanoA <> frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex) Then
        If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTbm_Franquia_Adicional, lMensagem) = False Then
            Call gM46V999.gpGraLog(2, lMensagem)
            Exit Sub
        End If

        lCodSequencia = gM46V999.gCodSequencia

NovaSequencia:

        lSelect = " Select distinct Perc_Franquia, Fator_Ajuste from Tbm_Franquia_Adicional "
        lSelect = lSelect & " Where Cod_Ramo = " & gM46V999.gCod_Ramo
        lSelect = lSelect & " and Cod_Modalidade = " & frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex)
        lSelect = lSelect & " and Cod_Tarif     = " & gM46V999.gCodProduto
        lSelect = lSelect & " and Cod_Sequencia <= " & lCodSequencia
        lSelect = lSelect & " order by  Perc_Franquia "

        If gM46V999.gfObtRegistro(lbdTbm_Franquia_Adicional, lSelect, lrsTbm_Franquia_Adicional, lMensagem) = False Then
            Call gM46V999.gpFecha2(lbdTbm_Franquia_Adicional)
            Call gM46V999.gpGraLog(2, lMensagem)
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        If lrsTbm_Franquia_Adicional.EOF Then
            If lCodSequencia > 0 Then
                lCodSequencia = lCodSequencia - 1
                GoTo NovaSequencia
            End If
        End If

        Set mColecaoAdicional = New Collection
        Dim lobjstcAux As stcAuxiliar
        Do While Not lrsTbm_Franquia_Adicional.EOF
            Set lobjstcAux = New stcAuxiliar
            lobjstcAux.Perc_Franquia = lrsTbm_Franquia_Adicional("Perc_Franquia")
            lobjstcAux.Fator_Ajuste = lrsTbm_Franquia_Adicional("Fator_Ajuste")
            mColecaoAdicional.Add lobjstcAux
            lrsTbm_Franquia_Adicional.MoveNext
        Loop

        mUltProdutoA = gM46V999.gCodProduto
        mUltPlanoA = frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex)

        Call gM46V999.gpFecha2(lbdTbm_Franquia_Adicional)
    End If

    If gM46V999.CDblx(flxSelecionadas.TextMatrix(flxSelecionadas.Row, ColValFranMin)) <> 0 Then
        If flxSelecionadas.TextMatrix(flxSelecionadas.Row, ColTipFranq) = "Dias" Then
            Aux = flxSelecionadas.TextMatrix(flxSelecionadas.Row, ColValFranMin)
        Else
        Aux = Format(flxSelecionadas.TextMatrix(flxSelecionadas.Row, ColValFranMin), "standard")
        End If
        Aux = Space(12 - Len(Aux)) & Aux
        LstFranquias.AddItem Aux & "            - 0"
    End If

    For Each lobjstcAux In mColecaoAdicional
        If flxSelecionadas.TextMatrix(flxSelecionadas.Row, ColTipFranq) = "Dias" Then
            Aux = lobjstcAux.Perc_Franquia * gM46V999.CDblx(flxSelecionadas.TextMatrix(flxSelecionadas.Row, ColValFranMin))
        Else
        Aux = Format(lobjstcAux.Perc_Franquia * gM46V999.CDblx(flxSelecionadas.TextMatrix(flxSelecionadas.Row, ColValFranMin)), "standard")
        End If
        
        lobjstcAux.Valor = Aux
        If gM46V999.CDblx(flxSelecionadas.TextMatrix(flxSelecionadas.Row, ColValFranMin)) < Aux Then
            If Len(Aux) <= 12 Then
                Aux = Space(12 - Len(Aux)) & Aux
            Else
                Aux = Aux
            End If
            lobjstcAux.Descricao = Aux & "                       - " & FormataDesconto(lobjstcAux.Fator_Ajuste)
            LstFranquias.AddItem lobjstcAux.Descricao
        End If
    Next lobjstcAux

    'Acesso Syas Privilégios
    If Not gM46V999.M46O001.PermiteAlterarFranquia = True Then
       LstFranquias.RemoveItem (0) '= 0,00  ou Branco
    End If

End Sub


Private Sub mpBuscaLimFranquiaDesativado(ByVal pCodCobert As Integer, ByRef PerMin As Double, ByRef ValFraqMin As Double, ByRef ValFator As Double)
'Desativado Simony
    Dim lbdTbm_Fat_Franquia As ADODB.Connection         'Arquivo: P00Multi.
    Dim lMensagem As String         'Mensagem.
    Dim lrsTbm_Fat_Franquia As clsYasRecordSet          'Registro : Tbm_Fator_Franquia.
    Dim lCodSequencia As Integer
    Dim lCod_Clasf As String
    Dim lCod_Compl_Clasf As String
    Dim lCod_Constr As Integer
    Dim lTip_Cobert As Integer
    Dim lCEP As String
    Dim lSelect As String

    If Me.Tag = "Desabilitado" Then Exit Sub

    'Ficha 1000305
    If gM46V999.gCod_Ramo = e_CodRamo_Empresarial And frmT46V101A.cboConstrucao.ListIndex > -1 Then
        Select Case frmT46V101A.cboConstrucao.ListIndex
        Case 1
            lCod_Constr = e_CodContr_Superior
        Case 2
            lCod_Constr = e_CodContr_Inferior
        Case 3
            lCod_Constr = e_CodContr_Mista
        Case Else
            lCod_Constr = e_CodContr_Superior
        End Select
    Else
        If frmT46V101A.optContrucao(1).Value = True Then
            lCod_Constr = e_CodContr_Inferior
        Else
            lCod_Constr = e_CodContr_Superior
        End If
    End If

    'Ficha 1000305
    '== Tipo Cobertura
    If gM46V999.gCod_Ramo = e_CodRamo_Empresarial And frmT46V101A.cboTipCobert.ListIndex > -1 Then
        Select Case frmT46V101A.cboTipCobert.ListIndex
        Case 0
            lTip_Cobert = e_TipCobert_PredioeConteudo
        Case 1
            lTip_Cobert = e_TipCobert_Predio
        Case 2
            lTip_Cobert = e_TipCobert_Conteudo
        Case Else
            lTip_Cobert = e_TipCobert_PredioeConteudo
        End Select
    Else
        lTip_Cobert = e_TipCobert_PredioeConteudo
    End If
    gM46V999.gTipCobertura = lTip_Cobert

    lCod_Clasf = Left(frmT46V101A.txtClassificacao.Tag, 4)
    lCod_Compl_Clasf = Mid(frmT46V101A.txtClassificacao.Tag, 5, 2)

    PerMin = 0
    ValFraqMin = 0
    ValFator = 0
    ' If lTip_Atividade = 0 Then
    lTip_Atividade = gTip_Atividade
    ' End If

    If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTbm_Fat_Franquia, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If

    lCodSequencia = gM46V999.gCodSequencia

NovaSequencia1:

    '1 - Procura Registro de exceção com a chave completa
    'Procura o valor do Val_Franq_Min e Perc_Franq
    lCEP = val(Trim(Replace(frmT46V101A.txtRisCEP.Text, "-", "")))

    lSelect = " Select Val_Franq_Min as Min , Perc_Franquia as Perc, Fator_Ajuste as Fator from Tbm_Fator_Franquia "
    lSelect = lSelect & " Where Cod_Ramo = " & gM46V999.gCod_Ramo
    lSelect = lSelect & " and Cod_Modalidade = " & frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex)
    lSelect = lSelect & " and Cod_Tip_Reg = 'E' "
    If gM46V999.gCodProduto >= 940 Then
        lSelect = lSelect & " and ((Cod_Clasf = '" & lCod_Clasf & "'"
        lSelect = lSelect & " and Cod_Compl_Clasf = '" & lCod_Compl_Clasf & "')"
        lSelect = lSelect & " or (Cod_Clasf = '0'))"
    Else
        lSelect = lSelect & " and Cod_Clasf = '" & lCod_Clasf & "'"
        lSelect = lSelect & " and Cod_Compl_Clasf = '" & lCod_Compl_Clasf & "'"
    End If
    lSelect = lSelect & " and Cod_Cobert = " & pCodCobert
    lSelect = lSelect & " and Cod_Construcao in (0 , " & lCod_Constr & ")"
    lSelect = lSelect & " and Cod_Tarif = " & gM46V999.gCodProduto
    lSelect = lSelect & " and Cod_Sequencia <= " & lCodSequencia

    If gM46V999.gCodProduto >= 940 Then
        lSelect = lSelect & " AND ((CEP_INICIO <= " & lCEP
        lSelect = lSelect & " AND CEP_FIM >= " & lCEP & " ) "
        lSelect = lSelect & " OR (CEP_INICIO =  0 ))"
        lSelect = lSelect & " order by Prioridade desc "
    End If

    If gM46V999.gfObtRegistro(lbdTbm_Fat_Franquia, lSelect, lrsTbm_Fat_Franquia, lMensagem) = False Then
        Call gM46V999.gpFecha2(lbdTbm_Fat_Franquia)
        Call gM46V999.gpGraLog(2, lMensagem)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    '2 - Se não encontrar a chave completa vai procurar o padrão
    If lrsTbm_Fat_Franquia.EOF Then
        If lCodSequencia > 0 Then
            lCodSequencia = lCodSequencia - 1
            GoTo NovaSequencia1
        Else
            GoTo ProcuraDefault
        End If
    End If

    If lrsTbm_Fat_Franquia("Min") = 0 Then
        GoTo ProcuraDefault
    End If

    PerMin = lrsTbm_Fat_Franquia("Perc")
    ValFraqMin = lrsTbm_Fat_Franquia("Min")
    ValFator = lrsTbm_Fat_Franquia("Fator")

    GoTo FimBusca

ProcuraDefault:

    '3 - Procura o registro somente com Ramo, Modalidade, Cobertura e TipoAtividade
    'Procura o valor do Val_Franq_Min, Perc_Franq

    lCodSequencia = gM46V999.gCodSequencia

NovaSequencia2:

    lSelect = " Select Val_Franq_Min as Min , Perc_Franquia as Perc, Fator_Ajuste as Fator  from Tbm_Fator_Franquia "
    lSelect = lSelect & " Where Cod_Ramo = " & gM46V999.gCod_Ramo
    lSelect = lSelect & " and Cod_Modalidade = " & frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex)
    lSelect = lSelect & " and Cod_Tip_Reg = 'D' "
    lSelect = lSelect & " and Cod_Cobert = " & pCodCobert
    lSelect = lSelect & " and Tip_Atividade in (0,3," & lTip_Atividade & ")"
    lSelect = lSelect & " and Cod_Tarif = " & gM46V999.gCodProduto
    lSelect = lSelect & " and Cod_Sequencia <= " & lCodSequencia
    lSelect = lSelect & " order by Prioridade desc "

    If gM46V999.gfObtRegistro(lbdTbm_Fat_Franquia, lSelect, lrsTbm_Fat_Franquia, lMensagem) = False Then
        Call gM46V999.gpFecha2(lbdTbm_Fat_Franquia)
        Call gM46V999.gpGraLog(2, lMensagem)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    If lrsTbm_Fat_Franquia.EOF Then
        If lCodSequencia > 0 Then
            lCodSequencia = lCodSequencia - 1
            GoTo NovaSequencia2
        Else
            GoTo FimBusca
        End If
    End If

    'Simony
    'If lrsTbm_Fat_Franquia(0) = 0 Then
    '    GoTo FimBusca
    'End If

    PerMin = lrsTbm_Fat_Franquia("Perc")
    ValFraqMin = lrsTbm_Fat_Franquia("Min")
    ValFator = lrsTbm_Fat_Franquia("Fator")

FimBusca:
    Call gM46V999.gpFecha2(lbdTbm_Fat_Franquia)

    '1001135 - Corretor Volks - Franquia Especial Concessionaria
    If gM46V999.gCodProduto >= 1002 Then
        If (InStr(1, "00008001/00008002/00008003/00008004/00008005/00008006/00008008/00008011/00008014/00008028", Format(frmT46V102A.txtCorretor(1).Tag, "00000000"), vbTextCompare) > 0) And _
           (val(lCod_Clasf) = 22 And val(lCod_Compl_Clasf) = 13) And _
           (InStr(1, "0038/0064/0066", Format(pCodCobert, "0000"), vbTextCompare) > 0) Then
            PerMin = 15
            ValFraqMin = 2000
        End If
    End If

End Sub

Private Sub mpGetLimFranquia(ByVal pCodCobert As Integer, ByRef PerMin As Double, ByRef ValFraqMin As Double, ByRef ValFator As Double, TextoFranquia As String, DescCobert As String, pNum_Meses_PI As Integer)
    Dim lMensagem As String         'Mensagem.
    Dim lCod_Clasf As String
    Dim lCod_Compl_Clasf As String
    Dim lCod_Constr As Integer
    Dim lCEP As String
    Dim Val_IS As Double
    Dim ValFraqMax As Double
    Dim TemVideObs As String
    Dim i As Integer

    If Me.Tag = "Desabilitado" Then Exit Sub

    'Ficha 1000305
    If gM46V999.gCod_Ramo = e_CodRamo_Empresarial And frmT46V101A.cboConstrucao.ListIndex > -1 Then
        Select Case frmT46V101A.cboConstrucao.ListIndex
        Case 1
            lCod_Constr = e_CodContr_Superior
        Case 2
            lCod_Constr = e_CodContr_Inferior
        Case 3
            lCod_Constr = e_CodContr_Mista
        Case Else
            lCod_Constr = e_CodContr_Superior
        End Select
    Else
        If frmT46V101A.optContrucao(1).Value = True Then
            lCod_Constr = e_CodContr_Inferior
        Else
            lCod_Constr = e_CodContr_Superior
        End If
    End If

    lCod_Clasf = Left(frmT46V101A.txtClassificacao.Tag, 4)
    lCod_Compl_Clasf = Mid(frmT46V101A.txtClassificacao.Tag, 5, 2)
    lCEP = val(Trim(Replace(frmT46V101A.txtRisCEP.Text, "-", "")))

    Val_IS = 0
    For i = 1 To frmT46V117A.flxSelecionadas.Rows
        flxSelecionadas.Row = i
        If frmT46V117A.flxSelecionadas.TextMatrix(i, ColCodCobert) = pCodCobert Then
            Val_IS = frmT46V117A.flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColValLMG)
            Exit For
        End If
    Next

'    Call gM46V999.mpNovoBuscaLimFranquiaAvancada(gM46V999.gCod_Ramo, frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex), gM46V999.gCodProduto, _
'                                     pCodCobert, lCod_Clasf, lCod_Compl_Clasf, _
'                                     lCod_Constr, lCEP, gM46V999.gCodSequencia, Val(frmT46V102A.txtCorretor(1).Tag), _
'                                                ValFraqMin, PerMin, ValFator, ValFraqMax, TextoFranquia, DescCobert, pNum_Meses_PI, pMuliploFraq, pDescMultFraq, _
'                                                Val_IS)
                                                
     Call gM46V999.mpNovoBuscaLimFranquia(gM46V999.gCod_Ramo, frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex), gM46V999.gCodProduto, _
                                          pCodCobert, lCod_Clasf, lCod_Compl_Clasf, lCod_Constr, lCEP, gM46V999.gCodSequencia, val(frmT46V102A.txtCorretor(1).Tag), _
                                          ValFraqMin, PerMin, ValFator, ValFraqMax, TextoFranquia, DescCobert, pNum_Meses_PI)
 
End Sub

Private Function mfBuscaPerIndenitario(ByVal pCodCobert As Integer) As Boolean
    Dim lbdTbm_Coef_PI As ADODB.Connection    'Arquivo: P00Multi.
    Dim lMensagem As String         'Mensagem.
    Dim lrsTbm_Coef_PI As clsYasRecordSet     'Registro : Tbm_Coef_PI.
    Dim lSelect As String
    mfBuscaPerIndenitario = False

    If Me.Tag = "Desabilitado" Then
        Exit Function
    End If

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

    Call gM46V999.gpFecha2(lbdTbm_Coef_PI)

End Function

Private Function mfBuscaLucrosCessantes(ByVal pCodCobert As Integer) As Boolean
    Dim lbdTbm_Cob_LucroCess As ADODB.Connection    'Arquivo: P00Multi.
    Dim lMensagem As String   'Mensagem.
    Dim lrsTbm_Cob_LucroCess As clsYasRecordSet     'Registro : Tbm_Cob_LucroCess.
    Dim lSelect As String
    mfBuscaLucrosCessantes = False

    If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTbm_Cob_LucroCess, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Function
    End If

    '1 - Procura pela Cobertura na Tabela de Coeficiente
    'Se encontrar é permitido Cadastrar

    lSelect = " Select *  from Tbm_Cob_LucroCess "
    lSelect = lSelect & " Where Cod_Ramo = " & gM46V999.gCod_Ramo
    lSelect = lSelect & " and Cod_Cobert = " & pCodCobert
    lSelect = lSelect & " and Cod_Tarif = " & gM46V999.gCodProduto
    lSelect = lSelect & " and Cod_Sequencia <= " & gM46V999.gCodSequencia
    lSelect = lSelect & " order by Cod_Sequencia desc "

    If gM46V999.gfObtRegistro(lbdTbm_Cob_LucroCess, lSelect, lrsTbm_Cob_LucroCess, lMensagem) = False Then
        Call gM46V999.gpFecha2(lbdTbm_Cob_LucroCess)
        Call gM46V999.gpGraLog(2, lMensagem)
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    '2 - Se encontrar permite cadastro do Valor de Risco de Lucro Cessantes
    If Not lrsTbm_Cob_LucroCess.EOF Then
        mfBuscaLucrosCessantes = True
    End If

    Call gM46V999.gpFecha2(lbdTbm_Cob_LucroCess)

End Function

Private Sub VerificaAlteracaoFranquia(ByVal lCod_Clasf As String, ByVal lCod_Compl_Clasf As String)

    If gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Then
        gM46O001.PermiteAlterarFranquia = False
    End If

    With flxSelecionadas
        If gM46V999.gCod_Corretor = 8001 Or gM46V999.gCod_Corretor = 8002 Or gM46V999.gCod_Corretor = 8003 Or gM46V999.gCod_Corretor = 8004 Or _
           gM46V999.gCod_Corretor = 8005 Or gM46V999.gCod_Corretor = 8006 Or gM46V999.gCod_Corretor = 8008 Or gM46V999.gCod_Corretor = 8011 Or _
           gM46V999.gCod_Corretor = 8014 Or gM46V999.gCod_Corretor = 8028 Then
            If Trim(lCod_Clasf) = gM46V999.gfBuscarDePara("022", 2) And lCod_Compl_Clasf = gM46V999.gfBuscarDePara("022", 2, "13") Then
                If IsNumeric(.TextMatrix(.Row, ColCodCobert)) Then
                    If .TextMatrix(.Row, ColCodCobert) = 38 Or .TextMatrix(.Row, ColCodCobert) = 66 Then
                        gM46O001.PermiteAlterarFranquia = True
                    End If
                End If
            Else
                gM46O001.PermiteAlterarFranquia = False
            End If
        End If
    End With

    If gM46V999.gTipoAcesso = DIGITACAO And gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then
        gM46O001.PermiteAlterarFranquia = True
    End If

    If gM46O001.PermiteAlterarFranquia = True Then
        gM46O001.gAlterarFranquia = True
    Else
        gM46O001.gAlterarFranquia = False
    End If

End Sub

Private Sub CBO_ordenaCob_Click()
'[FRH]->
    Select Case CBO_ordenaCob.ListIndex
    Case 0:
        flxCobertura.ColSel = 0
        flxCobertura.Col = 0
        flxCobertura.Sort = 1
    Case 1:
        flxCobertura.ColSel = 1
        flxCobertura.Col = 1
        flxCobertura.Sort = 1
    End Select

    SendKeys "{TAB}"
    '<-[FRH]
End Sub

Private Sub cmdAdicionar_Click()
    Dim lIn As Long             'Auxiliar - Linha inicial.
    Dim lFim As Long             'Auxiliar - Linha final.
    Dim lI As Long             'Contador.
    Dim lMensagem As String    'Mensagem.
    Dim lCod_Clasf As String
    Dim lCod_Compl_Clasf As String
    Dim lcol As Integer
    Dim lNomCobert As String
    Dim lCodEstip As Integer
    Dim i As Integer
    Dim c As Integer

    If Me.Tag = "Desabilitado" Then Exit Sub

    If IncluirCobertura = e_SimNao_Nao Then Exit Sub
    
    If frmT46V101A.objstcPedido.VAL_PR_Informado <> 0 Then
        If MsgBox("Este cálculo possui premio informado." & vbCrLf & "Adicionar a cobertura vai cancelar o prêmio informado de todas as coberturas." & vbCrLf & "Deseja continuar?", vbOKCancel, "Prêmio Informado") = vbCancel Then Exit Sub
        frmT46V101A.objstcPedido.VAL_PR_Informado = 0
        For i = 1 To frmT46V101A.objstcPedido.ColecaoPedLoc.Count
            For c = 1 To frmT46V101A.objstcPedido.ColecaoPedLoc(i).ColecaoPedCobert.Count
                frmT46V101A.objstcPedido.ColecaoPedLoc(i).ColecaoPedCobert(c).PremioInformado = 0
                frmT46V101A.objstcPedido.ColecaoPedLoc(i).ColecaoPedCobert(c).VAL_PR = frmT46V101A.objstcPedido.ColecaoPedLoc(i).ColecaoPedCobert(c).Val_Pr_Orig
                frmT46V101A.objstcPedido.ColecaoPedLoc(i).ColecaoPedCobert(c).Val_Pr_Orig = 0
            Next c
        Next i
    End If
    
    frmT46V101A.objstcPedido.Cod_Sit_Ped = e_SitPedido_SemCalculo
    
    'Ficha 1000757
    lCod_Clasf = Left(frmT46V101A.txtClassificacao.Tag, 4)
    lCod_Compl_Clasf = Mid(frmT46V101A.txtClassificacao.Tag, 5, 2)
    lCodEstip = frmT46V101A.objstcPedido.Cod_Estip

    Call gpLimPremios
    If mLinhaFinal <> 0 And mLinhaInicial <> 0 Then
        If mLinhaInicial > mLinhaFinal Then
            lIn = mLinhaFinal
            lFim = mLinhaInicial
        Else
            lIn = mLinhaInicial
            lFim = mLinhaFinal
        End If
        lI = lIn
        flxSelecionadas.Redraw = False
        Do
            If flxCobertura.Rows > 1 And flxCobertura.RowSel <> 0 Then
                If lCodEstip = 30 Then
                'AllRisks
                If flxCobertura.TextMatrix(lI, 0) = 1172 Then
                    frmT46V202A.Show vbModal, Me
                End If
                
                'Obras de Arte
                If flxCobertura.TextMatrix(lI, 0) = 3022 Then
                    frmT46V204A.Show vbModal, Me
                End If
                End If
                
                'Bicicleta
                If flxCobertura.TextMatrix(lI, 0) = 1215 Then
                    frmT46V196A.Show vbModal, Me
                End If
                With flxSelecionadas
                    .Rows = .Rows + 1
                    .Row = .Rows - 1
                    .TextMatrix(.Row, ColCodCobert) = flxCobertura.TextMatrix(lI, 0)

                    If gM46V111.gCod_Dizer <> 0 Then
                        .TextMatrix(.Row, ColAjusteEndosso) = "I"  'Marca inclusao de cobertura no endosso
                    End If
                    .TextMatrix(.Row, ColAlterouFranquia) = ""

                    'Ficha 1000757
                    Call VerificaAlteracaoFranquia(lCod_Clasf, lCod_Compl_Clasf)
                    
                    'Elevador
                    If .TextMatrix(.Row, ColCodCobert) = 1103 And gM46V999.gCod_Ramo = e_CodRamo_Condominio Then
                        Select Case frmT46V101A.cboElevador.ItemData(frmT46V101A.cboElevador.ListIndex)
                        Case 0
                            .TextMatrix(.Row, ColNomCobert) = flxCobertura.TextMatrix(lI, 1)
                        Case 1
                            .TextMatrix(.Row, ColNomCobert) = flxCobertura.TextMatrix(lI, 1) & " COM ELEVADOR"
                        Case 2
                            .TextMatrix(.Row, ColNomCobert) = flxCobertura.TextMatrix(lI, 1) & " SEM ELEVADOR"
                        End Select
                    Else
                        .TextMatrix(.Row, ColNomCobert) = flxCobertura.TextMatrix(lI, 1)
                    End If
                    
                    .TextMatrix(.Row, ColValLMG) = "0,00"
                    .Col = ColValLMG
                    .CellBackColor = &H80000018
                    .TextMatrix(.Row, ColLimiteMaximo) = 0
                    .TextMatrix(.Row, ColTabCobert) = flxCobertura.TextMatrix(lI, 3)
                    '.TextMatrix(.Row, ColPerFranq) = flxCobertura.TextMatrix(lI, 4)
                    '.TextMatrix(.Row, ColValFranq) = flxCobertura.TextMatrix(lI, 5)
                    'lPercent = 0
                    'lMinimo = 0
                    'Call mpBuscaLimFranquia(flxCobertura.TextMatrix(lI, 0), lPercent, lMinimo, lValFatorFraq)

                    .Col = ColDescComl
                    If .Row > 2 Then
                        .TextMatrix(.Row, ColDescComl) = .TextMatrix(.Row - 1, ColDescComl)
                    Else
                        .TextMatrix(.Row, ColDescComl) = "0,00"
                    End If
                    If .TextMatrix(.Row, ColDescComl) = "0,00" Then
                        .CellBackColor = &H80000018
                    Else
                        .CellBackColor = vbWhite
                    End If
                    .Col = ColDescExp
                    If .Row > 2 Then
                        .TextMatrix(.Row, ColDescExp) = .TextMatrix(.Row - 1, ColDescExp)
                    Else
                        .TextMatrix(.Row, ColDescExp) = "0,00"
                    End If
                    If .TextMatrix(.Row, ColDescExp) = "0,00" Then
                        .CellBackColor = &H80000018
                    Else
                        .CellBackColor = vbWhite
                    End If

                    .TextMatrix(.Row, ColCobObrigat) = flxCobertura.TextMatrix(lI, 6)
                    .TextMatrix(.Row, ColPerLimite) = Replace(flxCobertura.TextMatrix(lI, 2), "%", "")
                    .TextMatrix(.Row, ColCodGrupoCobert) = flxCobertura.TextMatrix(lI, 8)

                    If mfBuscaPerIndenitario(flxCobertura.TextMatrix(lI, 0)) Then
                        ' INICIO - Tarifa 1056 e 1057 - Buscar Período Indenitário Default
                        Dim intPeriodoIndenitarioDefault As Integer
                        intPeriodoIndenitarioDefault = mfBuscaPerIndenitarioDefault(flxCobertura.TextMatrix(lI, 0))
                        ' FIM - Tarifa 1056 e 1057 - Buscar Período Indenitário Default

                        .TextMatrix(.Row, ColPermitePInd) = True
                        .TextMatrix(.Row, ColPInd) = intPeriodoIndenitarioDefault
                        .Col = ColPInd
                        .CellBackColor = vbWhite
                    End If
                    ' Tarifa 1056 e 1057 - Inicio
                    If mfBuscaLucrosCessantes(flxCobertura.TextMatrix(lI, ColCodCobert)) Or mfBuscaCoberturasValorRiscoLMGAgravo(flxCobertura.TextMatrix(lI, ColCodCobert)) >= 0# Then
                        .Col = ColValRisco
                        .CellBackColor = vbWhite
                    End If
                    ' Tarifa 1056 e 1057 - Fim

                    .TextMatrix(.Row, ColCobertPerLimite) = flxCobertura.TextMatrix(lI, 7)
                    .TextMatrix(.Row, ColCodGrupoCobert) = flxCobertura.TextMatrix(lI, 8)
                    Call mpCalcularLimite(.Row)
                    .TextMatrix(.Row, ColValRisco) = "0,00"
                    mCobertSelec = mCobertSelec & .TextMatrix(.Row, ColCodCobert) & ";"
                    .TextMatrix(.Row, ColValISAte) = mfRetornaLMG(.Row)
                    .TextMatrix(.Row, ColValISmin) = mfRetornaLMGMin(.Row)

                    Call mpVerificarLimiteDependente
                    Call mpAtualizaFranquiaMin(.Row)
                End With
                If flxCobertura.Rows = 2 Then
                    flxCobertura.Rows = 1
                    flxCobertura.Row = 0
                Else
                    flxCobertura.RemoveItem lI
                End If
            End If
            lIn = lIn + 1
            If lIn > lFim Then Exit Do
        Loop
        flxCobertura.Row = 0
    End If

    'LMG(LC)
    Call mpRecalcularLMGLC
    flxSelecionadas.Redraw = True
    mLinhaFinal = 0
    mLinhaInicial = 0
    
    '20171221: Siscota End
    frmT46V101A.flagCalcular = True
    
End Sub

Public Sub mpCepAlterado()
    Dim lI As Integer
    For lI = 1 To flxSelecionadas.Rows - 1
        flxSelecionadas.TextMatrix(lI, ColAlterouFranquia) = ""
        Call mpAtualizaFranquiaMin(lI)
    Next
End Sub
Sub mpAtualizaFranquiaMin(pLin As Integer)
    Dim lPercent As Double
    Dim lMinimo As Double
    Dim lMinimoDias As Integer
    Dim lValFatorFraq As Double
    Dim lColAnt As Integer
    Dim lTextoFranquia As String
    Dim lDescFranquia As String
    Dim lRowAnt As Integer
    Dim pMuliploFraq As Integer
    Dim pDescMultFraq As Double
    Dim pobjStcCobert As stcA46V704B
    Dim lrsFranquia As clsYasRecordSet

    Set pobjStcCobert = New stcA46V704B

    If Me.Tag = "Desabilitado" Then Exit Sub
    If gM46V999.gCod_Ramo = e_CodRamo_Roubos Then Exit Sub
    'If gM46V999.gNovaPropostaLMI Then Exit Sub

    'Ficha 1000757
    'If gM46O001.PermiteAlterarFranquia Then Exit Sub


    If pLin >= flxSelecionadas.Rows Or pLin = 0 Then Exit Sub

    If gM46V111.gCod_Dizer <> 0 Then
        If flxSelecionadas.TextMatrix(pLin, ColAjusteEndosso) <> "I" Then   'Não é nova Cobertura
            Exit Sub
        End If
    End If

    If flxSelecionadas.TextMatrix(pLin, ColAlterouFranquia) = "SIM" And gM46V999.gCodProduto < 1080 Then Exit Sub   'Foi alterada Manualmente

    lPercent = 0
    lMinimo = 0
    lMinimoDias = 0
    lValFatorFraq = 0

    With flxSelecionadas
        lColAnt = .Col
        lRowAnt = .Row

        'Call mpBuscaLimFranquia(.TextMatrix(pLin, ColCodCobert), lPercent, lMinimo, lValFatorFraq)
        
        If gM46V999.gCodProduto >= 1080 Then
            Call gM46V999.mfFranqNovoLim(.TextMatrix(pLin, ColCodCobert), gM46V999.CDblx(.TextMatrix(pLin, ColValRisco)), frmT46V101A.objstcPedLoc.cod_clasf, frmT46V101A.objstcPedLoc.cod_compl_clasf, _
                                            frmT46V101A.objstcPedLoc.Cod_Constr, Replace(frmT46V101A.txtRisCEP, "-", ""), frmT46V101A.cboTipCobert.ItemData(frmT46V101A.cboTipCobert.ListIndex), _
                                            IIf(IsNumeric(.TextMatrix(pLin, ColPInd)), .TextMatrix(pLin, ColPInd), 0), lrsFranquia)
            If lrsFranquia.RecordCount > 0 Then
                lPercent = gM46V999.gfTratarNulo(lrsFranquia("Perc_Franquia"), tpoNumerico)
                lTextoFranquia = gM46V999.mfMontaTexofranquia(gM46V999.gfTratarNulo(lrsFranquia("Perc_Franquia"), tpoNumerico), gM46V999.gfTratarNulo(lrsFranquia("Val_Franq_Min"), tpoNumerico), gM46V999.gfTratarNulo(lrsFranquia("Dias_Franq_Min"), tpoNumerico), gM46V999.gfTratarNulo(lrsFranquia("Texto"), tpoString), gM46V999.gfTratarNulo(lrsFranquia("Compl"), tpoString), gM46V999.gfTratarNulo(lrsFranquia!Tem_VideObs, tpoString))
                lMinimo = gM46V999.gfTratarNulo(lrsFranquia("Val_Franq_Min"), tpoNumerico)
                lMinimoDias = gM46V999.gfTratarNulo(lrsFranquia("Dias_Franq_Min"), tpoNumerico)
                lValFatorFraq = gM46V999.gfTratarNulo(lrsFranquia("Fator_Ajuste"), tpoNumerico)
                
               ' If frmT46V101A.objstcPedido.Cod_Estip = 20 Then ' Imobiliário
                 '   lMinimo = gM46V999.gObtemValFranqProRata(frmT46V101A.objstcPedido.Dat_Ini_Vig, frmT46V101A.objstcPedido.Dat_Ter_Vig, lMinimo, False)
                    'lMinimoDias = gM46V999.gObtemValFranqProRata(frmT46V101A.objstcPedido.Dat_Ini_Vig, frmT46V101A.objstcPedido.Dat_Ter_Vig, lMinimoDias, True)
               ' End If
            End If
                
        Else
            Call mpGetLimFranquia(flxSelecionadas.TextMatrix(pLin, ColCodCobert), lPercent, lMinimo, lValFatorFraq, lTextoFranquia, lDescFranquia, val(flxSelecionadas.TextMatrix(pLin, ColPInd)))
        End If

        .TextMatrix(pLin, ColPerFranq) = lPercent
        .TextMatrix(pLin, ColTextoFranquia) = lTextoFranquia
        .TextMatrix(pLin, ColDescFranquia) = lDescFranquia
        
        If lMinimoDias = 0 And IIf(IsNumeric(.TextMatrix(pLin, ColPInd)), .TextMatrix(pLin, ColPInd), 0) = 0 Then
            .TextMatrix(pLin, ColTipFranq) = "R$"
        .TextMatrix(pLin, ColValFranq) = Format(lMinimo, "standard")
        .TextMatrix(pLin, ColValFranMin) = Format(lMinimo, "standard")
        Else
            .TextMatrix(pLin, ColTipFranq) = "Dias"
            .TextMatrix(pLin, ColValFranq) = lMinimoDias
            .TextMatrix(pLin, ColValFranMin) = lMinimoDias
        End If
        
        .TextMatrix(pLin, ColCoefFranq) = lValFatorFraq
        .TextMatrix(pLin, ColMultiploFranquia) = IIf(pMuliploFraq = 0, 1, pMuliploFraq) & "x"
        .TextMatrix(pLin, ColDesconto) = Format(pDescMultFraq, "standard")

        If gM46V999.gCodProduto >= 1080 Then
        Call frmT46V179A.mpIncluirFranquia(.TextMatrix(pLin, ColCodCobert), lDescFranquia, lTextoFranquia, pDescMultFraq, gM46V999.gfTratarNulo(lrsFranquia("Observacao_Franq"), tpoString))
        Else
            Call frmT46V179A.mpIncluirFranquia(.TextMatrix(pLin, ColCodCobert), lDescFranquia, lTextoFranquia, pDescMultFraq, Empty)
        End If

        'Verifico se existe franquias adicionais.
        LstFranquias.Clear
        LstFranquias.AddItem " "
        If flxSelecionadas.TextMatrix(pLin, ColCodCobert) = 1 Then
            Call mpMostraFranquiasBasica(val(flxSelecionadas.TextMatrix(pLin, ColValLMG)))
        Else
            Call mpMostraFranquiasAdicional(val(flxSelecionadas.TextMatrix(pLin, ColValLMG)))
        End If

        If LstFranquias.ListCount > 1 Then    'Se só tiver um registro na lista, então será movida o registro da lista
            'Ficha 1000757
            'CR93231 - Franquia (permitir que seja selecionado uma franquia menos que o padrão) a partir da 1084
            If gM46V999.CDblx(.TextMatrix(pLin, ColValFranq)) < .TextMatrix(pLin, ColValFranMin) And gM46V999.gCodProduto < 1084 Then
                If .TextMatrix(flxSelecionadas.Rows - 1, ColTipFranq) = "Dias" Then
                    .TextMatrix(pLin, ColValFranq) = lMinimo
                Else
                    .TextMatrix(pLin, ColValFranq) = Format(lMinimo, "standard")
                End If
                .TextMatrix(pLin, ColPerFranq) = CDbl(lPercent)
            End If

            If gM46V999.CDblx(.TextMatrix(pLin, ColPerFranq)) < lPercent Then
                .TextMatrix(pLin, ColPerFranq) = Format(lPercent, "standard")
            End If
        End If

        .Row = pLin

        'Ficha 1000183
        If gM46V999.gCod_Ramo = e_CodRamo_Roubos Or gM46V999.gNovaPropostaLMI Or gM46O001.PermiteAlterarFranquia Then
            .Col = ColPerFranq
            .CellBackColor = &H80000018
        End If

        .Col = ColValFranq


        If .TextMatrix(pLin, ColValFranq) = "0,00" Then
            .CellBackColor = &H80000018
        Else
            .CellBackColor = vbWhite
        End If

        .Col = ColPerFranq

        If .TextMatrix(pLin, ColPerFranq) = "0" Then
            .CellBackColor = &H80000018
        Else
            .CellBackColor = vbWhite
        End If

        'Ficha 1000183
        If gM46V999.gCod_Ramo = e_CodRamo_Roubos Or gM46V999.gNovaPropostaLMI Or gM46O001.PermiteAlterarFranquia Then
            flxSelecionadas.Col = ColPerFranq
            flxSelecionadas.CellBackColor = &H80000018
        End If

        .Row = lRowAnt
    End With

End Sub

Private Sub cmdBonus_Click()
    Dim lValor As Integer

    If Me.Tag = "Desabilitado" Then Exit Sub

    Me.Visible = True
    Call frmT46V130A.mpInstanciarForm(lValor, Me)
    txtDescExp.Text = Format(lValor, "standard")
    Call cmdMoverDescExp_Click
End Sub

Private Sub cmdFranquia_Click()
    If flxSelecionadas.Row > 0 Then
        flxSelecionadas.TextMatrix(flxSelecionadas.Row, ColTipFranq) = Empty
        flxSelecionadas.TextMatrix(flxSelecionadas.Row, ColValFranq) = 0
        flxSelecionadas.TextMatrix(flxSelecionadas.Row, ColAlterouFranquia) = ""
        flxSelecionadas_SelChange
        mpAtualizaFranquiaMin flxSelecionadas.Row
    End If
End Sub

Private Sub cmdMover_Click()
    Dim lI As Integer    'Utilizado no for...loop.

    If Me.Tag = "Desabilitado" Then Exit Sub

    For lI = 1 To flxSelecionadas.Rows - 1
        flxSelecionadas.TextMatrix(lI, ColDescComl) = txtDescComercial.Text
        flxSelecionadas.Row = lI
        flxSelecionadas.Col = ColDescComl
        flxSelecionadas.CellBackColor = vbWhite
    Next
End Sub

Private Sub cmdMoverDescExp_Click()
    Dim lI As Integer    'Utilizado no for...loop.
    If Me.Tag = "Desabilitado" Then Exit Sub

    For lI = 1 To flxSelecionadas.Rows - 1
        flxSelecionadas.TextMatrix(lI, ColDescExp) = txtDescExp.Text
        flxSelecionadas.Row = lI
        flxSelecionadas.Col = ColDescExp
        flxSelecionadas.CellBackColor = vbWhite
    Next
End Sub

Private Sub cmdOk_Click()
    Dim lI As Integer
    Dim lValor As Double
    '3. Inibir a tela

    If Me.Tag <> "Desabilitado" Then
        flxSelecionadas_SelChange
    End If

    lValor = 0
    If gM46V999.gCod_Ramo = e_CodRamo_Roubos Then
        For lI = 1 To flxSelecionadas.Rows - 1
            lValor = lValor + flxSelecionadas.TextMatrix(lI, ColValLMG)
        Next lI
        frmT46V101A.txtISBas.Text = Format(lValor, "standard")
    End If

    'LMG(LC)
    Call mpRecalcularLMGLC

    Call ValidaPI

    On Error Resume Next
    '    Me.Hide
    DoEvents
    Me.Visible = False
    DoEvents
End Sub

Public Sub cmdRetirar_Click()
    Dim lIn As Long                 'Auxiliar - Linha inicial.
    Dim lFim As Long                 'Auxiliar - Linha fim.
    Dim lI As Long                 'Contador.
    Dim lExcluir As Boolean      '
    Dim lobjCobertura As clsA46V704A    'objeto da classe clsA46V704A (TAB_PED_COBERT)
    Dim lobjstcCobertura As stcA46V704B    'objeto da classe de estrutura stcA46V704B (TAB_PED_COBERT)
    Dim lMensagem As String                 'Mensagem.
    Dim lCobertObrigatoria As Boolean
    Dim i As Integer
    Dim c As Integer

    If AlterarCobertura = e_SimNao_Nao Then Exit Sub

    If Me.Tag = "Desabilitado" Then Exit Sub
    If flxSelecionadas.TextMatrix(flxSelecionadas.Row, ColCodGrupoCobert) = "4" Then Exit Sub

    If frmT46V101A.objstcPedido.VAL_PR_Informado <> 0 Then
        If MsgBox("Este cálculo possui premio informado." & vbCrLf & "Remover a cobertura vai cancelar o prêmio informado de todas as coberturas." & vbCrLf & "Deseja continuar?", vbOKCancel, "Prêmio Informado") = vbCancel Then Exit Sub
        frmT46V101A.objstcPedido.VAL_PR_Informado = 0
        For i = 1 To frmT46V101A.objstcPedido.ColecaoPedLoc.Count
            For c = 1 To frmT46V101A.objstcPedido.ColecaoPedLoc(i).ColecaoPedCobert.Count
                frmT46V101A.objstcPedido.ColecaoPedLoc(i).ColecaoPedCobert(c).PremioInformado = 0
                frmT46V101A.objstcPedido.ColecaoPedLoc(i).ColecaoPedCobert(c).VAL_PR = frmT46V101A.objstcPedido.ColecaoPedLoc(i).ColecaoPedCobert(c).Val_Pr_Orig
                frmT46V101A.objstcPedido.ColecaoPedLoc(i).ColecaoPedCobert(c).Val_Pr_Orig = 0
            Next c
        Next i
    End If
    
    frmT46V101A.objstcPedido.Cod_Sit_Ped = e_SitPedido_SemCalculo

    Call gpLimPremios
    If mLinhaFinal <> 0 And mLinhaInicial <> 0 Then
        If mLinhaInicial > mLinhaFinal Then
            lIn = mLinhaFinal
            lFim = mLinhaInicial
        Else
            lIn = mLinhaInicial
            lFim = mLinhaFinal
        End If
        lI = lIn
        If flxSelecionadas.Rows <= lIn Then
            mLinhaFinal = 0
            mLinhaInicial = 0
            Exit Sub
        End If
        Do
            If frmT46V101A.cboTipCobert.ListIndex = e_TipCobert_PredioeConteudo And _
               frmT46V101A.optVerbasSeparadas(0).Value = True And _
               (InStr(flxSelecionadas.TextMatrix(lI, ColNomCobert), DESC_COBERT_PREDIO) > 0 Or InStr(flxSelecionadas.TextMatrix(lI, ColNomCobert), DESC_COBERT_CONTEUDO) > 0) Then
                lCobertObrigatoria = True
            Else
                lCobertObrigatoria = False
            End If

            If flxSelecionadas.Rows > 1 And flxSelecionadas.RowSel <> 0 And (val(flxSelecionadas.TextMatrix(lI, ColCodCobert)) <> 1 Or (val(flxSelecionadas.TextMatrix(lI, ColCodCobert)) = 1 And gM46V999.gCod_Ramo = e_CodRamo_Roubos)) And lCobertObrigatoria = False Then
                'If flxSelecionadas.Rows > 1 And flxSelecionadas.RowSel <> 0 And (Val(flxSelecionadas.TextMatrix(lI, ColCodCobert)) <> 1 Or (Val(flxSelecionadas.TextMatrix(lI, ColCodCobert)) = 1 And gM46V999.gCod_Ramo = e_CodRamo_Roubos)) Then
                'Dependendo do plano, as coberturas não poderão ser excluídas

                If flxSelecionadas.TextMatrix(lI, ColCobObrigat) = 0 Then
                    If gM46V111.gCod_Dizer <> 0 Then
                        If flxSelecionadas.TextMatrix(lI, ColAjusteEndosso) = "I" Then   'Nova Cobertura
                            GoTo PodeRemover
                        Else
                            Call gM46V999.gpGraLog(1, "AGE0047 - Não é possível excluir cobertura do endosso. LMG será zerada.")
                            flxSelecionadas.TextMatrix(lI, ColValLMG) = 0
                        End If
                    Else
PodeRemover:
                        If (frmT46V101A.cboTipCobert.ListIndex = e_TipCobert_PredioeConteudo) And _
                           (frmT46V101A.optVerbasSeparadas(0).Value = True) And _
                           (InStr(UCase(flxSelecionadas.TextMatrix(lI, ColNomCobert)), DESC_COBERT_CONTEUDO)) > 0 Then
                            'Prédio e conteúdo com verbas separadas.
                        Else
                            With flxCobertura
                                .Rows = .Rows + 1
                                .Row = .Rows - 1
                                .TextMatrix(.Row, 0) = flxSelecionadas.TextMatrix(lI, ColCodCobert)
                                .TextMatrix(.Row, 1) = flxSelecionadas.TextMatrix(lI, ColNomCobert)
                                .TextMatrix(.Row, 2) = flxSelecionadas.TextMatrix(lI, ColPerLimite) & "%"
                                .TextMatrix(.Row, 3) = flxSelecionadas.TextMatrix(lI, ColTabCobert)
                                .TextMatrix(.Row, 4) = flxSelecionadas.TextMatrix(lI, ColPerFranq)
                                .TextMatrix(.Row, 5) = flxSelecionadas.TextMatrix(lI, ColValFranq)
                                .TextMatrix(.Row, 6) = 0
                                .TextMatrix(.Row, 7) = flxSelecionadas.TextMatrix(lI, ColCobertPerLimite)
                                .TextMatrix(.Row, 8) = flxSelecionadas.TextMatrix(lI, ColCodGrupoCobert)
                                mCobertSelec = Replace(mCobertSelec, ";" & flxSelecionadas.TextMatrix(lI, ColCodCobert) & ";", ";")

                            End With
                            If flxSelecionadas.Rows = 2 Then
                                flxSelecionadas.Rows = 1
                                flxSelecionadas.Row = 0
                            Else
                                Dim lCodCobertExcluida As String
                                
                                Call frmT46V179A.mpExcluirFranquia(flxSelecionadas.TextMatrix(lI, ColCodCobert))
                                flxSelecionadas.RemoveItem lI
                        
                            End If
                        End If
                    End If
                    Call mpVerificarLimiteDependente
                Else
                    Call gM46V999.gpGraLog(1, "AGE0047 - Não é possível excluir. Cobertura obrigatória.")
                End If
            End If
            lIn = lIn + 1
            If lIn > lFim Then Exit Do
        Loop
        flxSelecionadas.Row = 0
    End If

    'LMG(LC)
    Call mpRecalcularLMGLC

    mLinhaFinal = 0
    mLinhaInicial = 0
End Sub

Private Sub cmdTexto_Click()
    flxSelecionadas_SelChange
    Call frmT46V179A.mpAtualizarFranquias(frmT46V101A.objstcPedLoc.ColecaoPedCobert)
    frmT46V179A.Show 1
End Sub

Private Sub flxCobertura_DblClick()

    If Me.Tag = "Desabilitado" Then Exit Sub

    If flxCobertura.MouseRow = 0 Then

        If lLastColSort <> flxCobertura.Col Then
            lSortCriteria = 1
        Else
            lSortCriteria = IIf(lSortCriteria = 1, 2, 1)
        End If

        flxCobertura.ColSel = flxCobertura.Col
        lLastColSort = flxCobertura.Col
        flxCobertura.Sort = lSortCriteria
        If flxCobertura.Rows = 1 Then
            Call cmdAdicionar_Click
        End If
    Else
        Call cmdAdicionar_Click
    End If
End Sub

Private Sub flxCobertura_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mLinhaInicial = flxCobertura.RowSel
End Sub

Private Sub flxCobertura_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    mLinhaFinal = flxCobertura.RowSel
End Sub

Private Sub flxCobertura_Scroll()
    With flxCobertura
        .ColWidth(3) = 0
        .ColWidth(4) = 0
        .ColWidth(5) = 0
        .ColWidth(6) = 0
    End With
End Sub

Private Sub flxCobertura_SelChange()
    With flxCobertura
        .ColWidth(3) = 0
        .ColWidth(4) = 0
        .ColWidth(5) = 0
        .ColWidth(6) = 0
    End With
End Sub

Private Sub flxSelecionadas_Click()
    Dim lCod_Clasf As String
    Dim lCod_Compl_Clasf As String

    If Me.Tag = "Desabilitado" Then Exit Sub
    If flxSelecionadas.TextMatrix(flxSelecionadas.Row, ColCodGrupoCobert) = "4" Then Exit Sub

    mLinhaFocus = flxSelecionadas.RowSel
    If mLinhaFocus = 0 Then Exit Sub

    'Ficha 1000183
    'Ficha 1000304
    '' If flxSelecionadas.ColSel = ColValFranq And gM46V999.gCod_Ramo <> e_CodRamo_Roubos And Not gM46V999.gNovaPropostaLMI And Not gM46O001.PermiteAlterarFranquia Then

    'Ficha 1000757
    lCod_Clasf = Left(frmT46V101A.txtClassificacao.Tag, 4)
    lCod_Compl_Clasf = Mid(frmT46V101A.txtClassificacao.Tag, 5, 2)


    If flxSelecionadas.ColSel = ColValFranq And gM46V999.gCod_Ramo <> e_CodRamo_Roubos And Not gM46V999.gNovaPropostaLMI Then
        LstFranquias.Clear
        LstFranquias.AddItem " "

        'Ficha 1000757
        Call VerificaAlteracaoFranquia(lCod_Clasf, lCod_Compl_Clasf)

        If flxSelecionadas.TextMatrix(flxSelecionadas.Row, ColCodCobert) = 1 Then
            Call mpMostraFranquiasBasica(flxSelecionadas.TextMatrix(flxSelecionadas.Row, ColValLMG))
        Else
            Call mpMostraFranquiasAdicional(gM46V999.CDblx(flxSelecionadas.TextMatrix(flxSelecionadas.Row, ColValLMG)))
        End If
        LstFranquias.Left = flxSelecionadas.CellLeft + 50
        LstFranquias.Width = flxSelecionadas.CellWidth
        LstFranquias.Top = flxSelecionadas.Top + flxSelecionadas.CellTop + 220
        LstFranquias.Visible = True
        LstFranquias.SetFocus
    Else
        Call DesableFranquia
    End If

    If flxSelecionadas.ColSel = ColMultiploFranquia And gM46V999.gCod_Ramo <> e_CodRamo_Roubos And Not gM46V999.gNovaPropostaLMI Then

        'Ficha 1000757
        Call VerificaAlteracaoFranquia(lCod_Clasf, lCod_Compl_Clasf)

        If flxSelecionadas.TextMatrix(flxSelecionadas.Row, ColCodCobert) = 1 Then
            Call mpMostraFranquiasBasica(gM46V999.CDblx(flxSelecionadas.TextMatrix(flxSelecionadas.Row, ColValLMG)))
        Else
            Call mpMostraFranquiasAdicional(gM46V999.CDblx(flxSelecionadas.TextMatrix(flxSelecionadas.Row, ColValLMG)))
        End If
    End If
End Sub

Private Sub flxSelecionadas_GotFocus()
    flxSelecionadas.Redraw = True
End Sub

Private Sub flxSelecionadas_LeaveCell()
    If gCarga = False And TrocandoCor = False Then
        Call mpVerificarLimiteDependente
    End If
End Sub

Private Sub flxSelecionadas_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 2 And flxSelecionadas.ColSel = ColValFranq Then
        flxSelecionadas_Click
    Else
        mLinhaInicial = flxSelecionadas.RowSel
        flxSelecionadas.ToolTipText = flxSelecionadas.TextMatrix(flxSelecionadas.RowSel, ColNomCobert)
    End If
    
End Sub

Private Sub flxSelecionadas_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    mLinhaFinal = flxSelecionadas.RowSel
End Sub

Private Sub flxSelecionadas_DblClick()

    If Me.Tag = "Desabilitado" Then Exit Sub


    If flxSelecionadas.MouseRow = 0 Then
        If lLastColSortSel <> flxSelecionadas.Col Then
            lSortCriteriaSel = 1
        Else
            lSortCriteriaSel = IIf(lSortCriteriaSel = 1, 2, 1)
        End If

        flxSelecionadas.ColSel = flxSelecionadas.Col
        lLastColSortSel = flxSelecionadas.Col
        flxSelecionadas.Sort = lSortCriteriaSel
        If flxSelecionadas.Rows = 1 Then
            Call cmdRetirar_Click
        End If
    Else
        Call cmdRetirar_Click
    End If

End Sub

Private Sub flxSelecionadas_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lCod_Clasf As String
    Dim lCod_Compl_Clasf As String

    If Me.Tag = "Desabilitado" Or gM46V999.gEplano Or gM46V999.gEplanoSimplificado Then Exit Sub
    If flxSelecionadas.TextMatrix(flxSelecionadas.Row, ColCodGrupoCobert) = "4" Then Exit Sub

    'Ficha 1000757
    lCod_Clasf = Left(frmT46V101A.txtClassificacao.Tag, 4)
    lCod_Compl_Clasf = Mid(frmT46V101A.txtClassificacao.Tag, 5, 2)


    With flxSelecionadas
        .Row = .RowSel

        'Ficha 1000757
        Call VerificaAlteracaoFranquia(lCod_Clasf, lCod_Compl_Clasf)

        If .Col = ColPInd And .TextMatrix(.RowSel, ColPermitePInd) <> "True" Then Exit Sub
        If .RowSel > 0 Then
            Select Case .ColSel
            Case ColValLMG
                If .TextMatrix(.RowSel, ColCobObrigat) = 1 Then Exit Sub
            Case ColDescComl
                If .TextMatrix(.RowSel, ColCobObrigat) = 1 And _
                   frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex) > 2000 Then
                    Exit Sub
                End If
                If gM46V999.PerfilAcesso = e_Perfil_CALC_Gestor Then
                    If frmT46V101A.chkLMI.Value = vbChecked And gM46V999.gCodProduto >= 988 Then
                        'Não pode editar o desconto quando for LMI
                        Exit Sub
                    End If
                End If
            Case ColTaxaFinal
                If Not gM46O001.PermiteAlterarTaxaFinal Then Exit Sub
                
            Case ColValFranq
                
            'Acesso Syas Privilégios
            If Not gM46V999.M46O001.PermiteAlterarFranquia = True Then
                Exit Sub
            End If
            
            Case ColPerFranq
                'Ficha 1000183
                If gM46V999.gCod_Ramo <> e_CodRamo_Roubos And Not gM46V999.gNovaPropostaLMI And Not gM46O001.PermiteAlterarFranquia Then Exit Sub
                'Case ColValLMG, ColDescComl, ColTaxaFinal, ColValRisco - thais
            Case ColValLMG, ColValFranq, ColDescComl, ColPInd, ColTaxaFinal, ColValRisco
                'COLUNAS PERMITIDAS
            Case Else
                Exit Sub
            End Select
            Select Case KeyCode
                'Ficha 1000183
            Case 48 To 57, 97, 101
                If mLinhaFocus = .RowSel Then
                    mLinhaFocus = 0
                    mColunaFocus = 0
                    .TextMatrix(.RowSel, .ColSel) = ""
                    If .ColSel = ColPInd Then
                        RowIsAlterada = .RowSel
                    End If
                    .Row = .RowSel
                    .Col = .ColSel
                    .CellBackColor = vbWhite
                    .Refresh
                End If

            Case 46
                .Row = .RowSel
                .Col = .ColSel
                If .ColSel = ColValRisco And .CellBackColor <> vbWhite And .CellBackColor <> &H80000018 Then
                    Exit Sub
                End If
                .CellBackColor = &H80000018
                .Refresh
                'Ficha 1000183
                If .ColSel = ColPInd Then    'And Not gM46O001.PermiteAlterarFranquia Then - thais
                    .TextMatrix(.RowSel, .ColSel) = ""
                Else
                    If .ColSel = ColValLMG Then
                        'If AlterarIs = e_SimNao_Sim Then
                        .TextMatrix(.RowSel, .ColSel) = "0,00"
                        RowIsAlterada = .RowSel
                        'End If
                    Else
                        .TextMatrix(.RowSel, .ColSel) = "0,00"
                        .TextMatrix(.RowSel, ColTaxaFinal) = "0,0000"
                    End If
                End If
                If .ColSel = ColValLMG And .TextMatrix(.Row, ColCodCobert) = 1 And frmT46V101A.txtISBas.Enabled Then
                    frmT46V101A.txtISBas.Text = "0,00"
                End If

                'LMG(LC)
                If .ColSel = ColValLMG And .TextMatrix(.Row, ColCodCobert) = 3 Then
                    frmT46V101A.txtLucros_Cessantes.Text = "0,00"
                End If
            Case 32, 93
                flxSelecionadas_Click
            End Select
        End If
    End With
    Call gpHabSalvar
End Sub

Private Sub flxSelecionadas_KeyPress(KeyAscii As Integer)
    Dim lCod_Clasf As String
    Dim lCod_Compl_Clasf As String
    Dim lCol_ant As Integer
    Dim lLin As Integer
    Dim lVlrPredioConteudo As Double
    Dim lCobertObrigatoria As Boolean
    Dim lFranqDias As Integer

    If Me.Tag = "Desabilitado" Or gM46V999.gEplano Or gM46V999.gEplanoSimplificado Then Exit Sub
    If flxSelecionadas.TextMatrix(flxSelecionadas.Row, ColCodGrupoCobert) = "4" Then Exit Sub

    'Ficha 1000757
    lCod_Clasf = Left(frmT46V101A.txtClassificacao.Tag, 4)
    lCod_Compl_Clasf = Mid(frmT46V101A.txtClassificacao.Tag, 5, 2)


    With flxSelecionadas
        lCol_ant = .ColSel
        .Redraw = False
        DoEvents

        If .RowSel = 0 Then Exit Sub
        .Col = .ColSel

        'Ficha 1000757
        Call VerificaAlteracaoFranquia(lCod_Clasf, lCod_Compl_Clasf)

        Select Case .ColSel
        Case ColDescComl
            'coluna liberada
            If gM46V999.PerfilAcesso = e_Perfil_CALC_Gestor Then
                If frmT46V101A.chkLMI.Value = vbChecked And gM46V999.gCodProduto >= 988 Then
                    'Não pode editar o desconto quando for LMI
                    Exit Sub
                End If
            End If
        Case ColValRisco
            If .CellBackColor <> vbWhite And .CellBackColor <> &H80000018 Then
                Exit Sub
            End If
        Case ColValLMG
            If .TextMatrix(.RowSel, ColCobObrigat) = 1 Then
                Call gM46V999.gpGraLog(1, "AGE0048 - Calcular o prêmio, coberturas com valores definidos - F3.")
                Exit Sub
            End If
        Case ColTaxaFinal
            If Not gM46O001.PermiteAlterarTaxaFinal Then
                Exit Sub
            End If
            'Ficha 1000183
        Case ColPInd
            'If Not gM46O001.PermiteAlterarFranquia Then 'thais
            '    Exit Sub
            'Else
            '    .TextMatrix(.RowSel, ColPermitePInd) = "True"
            'End If
            If .TextMatrix(.RowSel, ColPermitePInd) <> "True" Then Exit Sub
            RowAlterouFranquia = .RowSel
            '.TextMatrix(.RowSel, ColTextoFranquia) = gM46V999.gfTextoFranquia(.TextMatrix(.RowSel, ColValFranq), .TextMatrix(.RowSel, ColPerFranq), Mid(.TextMatrix(.RowSel, ColTextoFranquia), 1, 1), .TextMatrix(.RowSel, ColCodCobert), Val(.TextMatrix(.RowSel, ColPInd)))
        Case ColPerFranq, ColValFranq
            'Ficha 1000183
'            If gM46V999.gCodProduto < 1080 And .ColSel = ColValFranq And Not IsNumeric(.TextMatrix(.RowSel, ColValFranq) & Chr(KeyAscii)) Then
'                 .Redraw = True
'                .TextMatrix(.RowSel, ColValFranq) = Format(0, "standard")
'                 .Redraw = False
'            End If
            If gM46V999.gCod_Ramo <> e_CodRamo_Roubos And Not gM46V999.gNovaPropostaLMI And Not gM46O001.PermiteAlterarFranquia Then Exit Sub
            
            If .ColSel = ColPerFranq And gM46V999.gCodProduto >= 1080 And .TextMatrix(.RowSel, ColTipFranq) = "Dias" Then
                .TextMatrix(.RowSel, .ColSel) = 0
                Exit Sub
            Else
                If Len(.TextMatrix(.RowSel, .ColSel)) > 1 And .TextMatrix(.RowSel, ColTipFranq) = "Dias" And KeyAscii <> 8 Then
                    KeyAscii = 0
                    Exit Sub
                End If
            End If
            RowAlterouFranquia = .RowSel
        
        Case ColDesconto
            RowAlterouFranquia = 0
            
        Case Else
            Exit Sub
        End Select

        Select Case KeyAscii
        Case 8, 44, 48 To 57
            If mLinhaFocus = .RowSel Then
                mLinhaFocus = 0
                mColunaFocus = 0
                .TextMatrix(.RowSel, .ColSel) = ""
                If .ColSel = ColValLMG Then
                    RowIsAlterada = .RowSel
                End If
                .Row = .RowSel
                .Col = .ColSel
                .CellBackColor = vbWhite
                .Refresh
            End If
        End Select
        Select Case KeyAscii
        Case 8
            If Len(.TextMatrix(.RowSel, .ColSel)) > 0 Then
                .TextMatrix(.RowSel, .ColSel) = Left(.TextMatrix(.RowSel, .ColSel), Len(.TextMatrix(.RowSel, .ColSel)) - 1)
            End If
            If Len(.TextMatrix(.RowSel, .ColSel)) = 0 Then
                .Col = .ColSel
                .CellBackColor = &H80000018
                If .ColSel = ColValLMG Then
                    RowIsAlterada = .RowSel
                End If
            End If
            If .ColSel = ColTaxaFinal Then
                If gM46V999.CDblx(.TextMatrix(.RowSel, .ColSel)) = 0 Then
                    .TextMatrix(.RowSel, ColTaxaAlterada) = ""
                End If
            End If

        Case 13
            If gM46V999.gCodProduto < 1080 Then
            
                If IsNumeric(.TextMatrix(.RowSel, .ColSel)) And .ColSel <> ColPInd And .ColSel <> ColTaxaFinal Then
                    flxSelecionadas.TextMatrix(.RowSel, .ColSel) = Format(.TextMatrix(.RowSel, .ColSel), "standard")
                End If
                If .RowSel < .Rows - 1 Then
                    .RowSel = .RowSel + 1
                    .Row = .RowSel
                    mLinhaFocus = flxSelecionadas.RowSel
                End If
                
            Else
            If IsNumeric(.TextMatrix(.RowSel, .ColSel)) And .ColSel <> ColPInd And .ColSel <> ColTaxaFinal Then
                If .ColSel = ColValFranq And .TextMatrix(.RowSel, ColTipFranq) = "Dias" Then
                    flxSelecionadas.TextMatrix(.RowSel, .ColSel) = .TextMatrix(.RowSel, .ColSel)
                Else
                flxSelecionadas.TextMatrix(.RowSel, .ColSel) = Format(.TextMatrix(.RowSel, .ColSel), "standard")
            End If
            End If
            
            If RowAlterouFranquia <> 0 Then
                Dim lI As Integer
                lI = RowAlterouFranquia
                'lFranqDias = IIf(.TextMatrix(lI, ColTipFranq) = "Dias", .TextMatrix(lI, ColValFranq), 0)
                lFranqDias = IIf(IsNumeric(.TextMatrix(lI, ColPInd)), .TextMatrix(lI, ColPInd), 0)
                .TextMatrix(lI, ColValFranq) = gM46V999.CDblx(.TextMatrix(lI, ColValFranq))
                .TextMatrix(lI, ColPerFranq) = gM46V999.CDblx(.TextMatrix(lI, ColPerFranq))
                .TextMatrix(lI, ColTextoFranquia) = gM46V999.gfTextoFranquia(.TextMatrix(lI, ColValFranq), .TextMatrix(lI, ColPerFranq), _
                                                    Mid(.TextMatrix(lI, ColTextoFranquia), 1, 1), .TextMatrix(lI, ColCodCobert), _
                                                    lFranqDias, .TextMatrix(lI, ColValRisco), frmT46V101A.objstcPedLoc)
                
                blnRECALCULA_MULTIPLO = False
                
                Dim Aux As String
                Dim i As Integer
                
                i = 1
                Do While i < LstFranquias.ListCount
                    Aux = Trim(LstFranquias.List(i))
                    If CDbl(flxSelecionadas.TextMatrix(lI, ColValFranq)) <= Trim(Mid(Aux, 1, InStr(Aux, "-") - 2)) Then
                        Exit Do
                    End If
                    i = i + 1
                Loop
                
                Aux = IIf(Trim(Mid(Aux, InStr(Aux, "-") + 2)) = Empty, 0, Trim(Mid(Aux, InStr(Aux, "-") + 2)))
                flxSelecionadas.TextMatrix(lI, ColDesconto) = Format(Aux, "standard")
                flxSelecionadas.TextMatrix(lI, ColCoefFranq) = 0
                Aux = IIf(i > 1, "3x", "1x")
                flxSelecionadas.TextMatrix(lI, ColMultiploFranquia) = Aux
                
                'CR93231 - Franquia (permitir que seja selecionado uma franquia menos que o padrão) a partir da 1084
                If gM46V999.CDblx(flxSelecionadas.TextMatrix(lI, ColValFranq)) < gM46V999.CDblx(flxSelecionadas.TextMatrix(lI, ColValFranMin)) And gM46V999.gCodProduto < 1084 Then
                    flxSelecionadas.TextMatrix(lI, ColValFranq) = flxSelecionadas.TextMatrix(lI, ColValFranMin)
                End If
            
                If flxSelecionadas.TextMatrix(lI, ColValFranMin) <> flxSelecionadas.TextMatrix(lI, ColValFranq) Then
                    flxSelecionadas.TextMatrix(lI, ColAlterouFranquia) = "SIM"
                Else
                    flxSelecionadas.TextMatrix(lI, ColAlterouFranquia) = Empty
                End If
            End If

            If .RowSel < .Rows - 1 Then
                .RowSel = .RowSel + 1
                .Row = .RowSel
                mLinhaFocus = flxSelecionadas.RowSel
            End If
            End If
            
        Case 44
            If InStr(1, .TextMatrix(.RowSel, .ColSel), Chr(44)) > 0 Then Exit Sub
            .TextMatrix(.RowSel, .ColSel) = .TextMatrix(.RowSel, .ColSel) & Chr(KeyAscii)
            .Col = .ColSel
            .CellBackColor = vbWhite
            If .ColSel = ColValLMG And .TextMatrix(.Row, 0) = 1 Then
                frmT46V101A.txtISBas.Text = Format(.TextMatrix(.RowSel, .ColSel), "standard")
                Call mpCalcularLimite(0)
                '  .TextMatrix(.RowSel, ColValRisco) = .TextMatrix(.RowSel, ColLimiteMaximo)
            Else
                'LMG(LC)
                If .ColSel = ColValLMG And .TextMatrix(.Row, 8) = 3 Then
                    frmT46V101A.txtLucros_Cessantes.Text = Format(.TextMatrix(.RowSel, .ColSel), "standard")
                End If
            End If
        Case 48 To 57
            If .ColSel = ColValLMG And .RowSel <> 1 Then
                If val(Replace(.TextMatrix(.RowSel, ColValLMG) & Chr(KeyAscii), ",", ".")) < 1000000000# Or gM46V999.gCod_Ramo = e_CodRamo_Roubos Then
                    .TextMatrix(.RowSel, .ColSel) = .TextMatrix(.RowSel, .ColSel) & Chr(KeyAscii)
                    RowIsAlterada = .RowSel
                End If
            Else
                If .ColSel = ColValFranq And .RowSel <> 1 Then
                    If Len(.TextMatrix(.RowSel, ColValFranq)) < (Len(.TextMatrix(.RowSel, ColLimiteMaximo)) - 1) Then
                        .TextMatrix(.RowSel, .ColSel) = .TextMatrix(.RowSel, .ColSel) & Chr(KeyAscii)
                        RowAlterouFranquia = .RowSel
                        'blnRECALCULA_MULTIPLO = True
                    End If
                Else
                    If .ColSel = ColDesconto And .RowSel <> 1 Then
                        If val(Replace(.TextMatrix(.RowSel, .ColSel) & Chr(KeyAscii), ",", ".")) <= 50# Then
                            .TextMatrix(.RowSel, .ColSel) = .TextMatrix(.RowSel, .ColSel) & Chr(KeyAscii)
                            RowAlterouFranquia = .RowSel
                            'blnRECALCULA_MULTIPLO = True
                        End If
                    Else
                        If .ColSel = ColTaxaFinal And .RowSel <> 1 Then
                            If val(Replace(.TextMatrix(.RowSel, ColTaxaFinal) & Chr(KeyAscii), ",", ".")) < 100# Then
                                .TextMatrix(.RowSel, .ColSel) = .TextMatrix(.RowSel, .ColSel) & Chr(KeyAscii)
                            End If
                        Else
                        .TextMatrix(.RowSel, .ColSel) = .TextMatrix(.RowSel, .ColSel) & Chr(KeyAscii)
                    End If
                End If
            End If
            End If
            .Col = .ColSel
            .CellBackColor = vbWhite
            If .ColSel = ColDescComl Then
                If .TextMatrix(.RowSel, .ColSel) > 99 Then .TextMatrix(.RowSel, .ColSel) = 99
            End If
            'Ficha 1000183
            If .ColSel = ColPInd Then   'thais
                If .TextMatrix(.RowSel, .ColSel) > 12 Then
                    'If Not gM46O001.PermiteAlterarFranquia Then
                    '.TextMatrix(.RowSel, .ColSel) = 12
                    .TextMatrix(.RowSel, ColPInd) = mfBuscaPerIndenitarioDefault(.TextMatrix(.RowSel, ColCodCobert))
                    'End If
                End If
                RowAlterouFranquia = .RowSel
            End If

            If frmT46V101A.cboTipCobert.ListIndex = e_TipCobert_PredioeConteudo And _
               frmT46V101A.optVerbasSeparadas(0).Value = True And _
               (InStr(.TextMatrix(.Row, ColNomCobert), DESC_COBERT_PREDIO) > 0 Or InStr(.TextMatrix(.Row, ColNomCobert), DESC_COBERT_CONTEUDO) > 0) _
               Then
                lCobertObrigatoria = True
            Else
                lCobertObrigatoria = False
            End If

            'If .ColSel = ColValLMG And .TextMatrix(.Row, ColCodCobert) = 1 Then
            'rodrigoS
            If .ColSel = ColValLMG And (.TextMatrix(.Row, ColCodCobert) = 1 Or lCobertObrigatoria = True) Then
                If (frmT46V101A.cboTipCobert.ListIndex = e_TipCobert_PredioeConteudo) And (frmT46V101A.optVerbasSeparadas(0).Value = True) Then
                    lVlrPredioConteudo = 0
                    For lLin = 1 To .Rows - 1
                        If lLin <> .RowSel Then
                            If InStr(.TextMatrix(lLin, ColNomCobert), DESC_COBERT_PREDIO) > 0 Or InStr(.TextMatrix(lLin, ColNomCobert), DESC_COBERT_CONTEUDO) > 0 Then
                                lVlrPredioConteudo = lVlrPredioConteudo + .TextMatrix(lLin, .ColSel)
                            End If
                        End If
                    Next
                    frmT46V101A.txtISBas.Text = Format(.TextMatrix(.RowSel, .ColSel) + lVlrPredioConteudo, "standard")
                Else
                    frmT46V101A.txtISBas.Text = Format(.TextMatrix(.RowSel, .ColSel), "standard")
                End If
                Call mpCalcularLimite(0)
                '     .TextMatrix(.RowSel, ColValRisco) = .TextMatrix(.RowSel, ColLimiteMaximo)
            End If

            'Correção Coef. VR X LMG
            'Email: ERRO SYAS COEFICIENTE VR/IS
            'Enviado: segunda-feira, 7 de novembro de 2011 18:12
            If .ColSel = ColValRisco And .TextMatrix(.Row, ColCodCobert) = 1 Then
                frmT46V101A.txtValoRisco.Text = Format(.TextMatrix(.RowSel, .ColSel), "standard")
            End If

            If .ColSel = ColTaxaFinal Then
                .TextMatrix(.RowSel, ColTaxaAlterada) = "Alterado"
                .CellForeColor = vbRed
            End If
            If .ColSel = ColValFranq Then
                .TextMatrix(.RowSel, ColAlterouFranquia) = "SIM"
                .Col = ColAlterouFranquia
                .CellForeColor = vbRed
                .Col = lCol_ant
                RowAlterouFranquia = .RowSel
            End If

            If .ColSel = ColPerFranq Then
                If CDbl(.TextMatrix(.RowSel, ColPerFranq)) > 99 Then
                    .TextMatrix(.RowSel, ColPerFranq) = 0
                Else
                    .TextMatrix(.RowSel, ColAlterouFranquia) = "SIM"
                    .Col = ColAlterouFranquia
                    .CellForeColor = vbRed
                    .Col = lCol_ant
                End If
                RowAlterouFranquia = .RowSel
            End If

            Dim lCoefFranq As Double
            
            If gM46V999.gCodProduto < 1080 Then
            If .ColSel = ColValFranq Then
                'thais
                If gM46V999.CDblx(.TextMatrix(.RowSel, ColValFranq)) <> gM46V999.CDblx(.TextMatrix(.RowSel, ColValFranMin)) Then
                    .TextMatrix(.RowSel, ColAlterouFranquia) = "SIM"
                    If .TextMatrix(.RowSel, ColCodCobert) = "1" Then
                        lCoefFranq = mfBuscaCoefFranquiasBasica(.TextMatrix(.RowSel, ColValLMG), .TextMatrix(.RowSel, ColValFranq))
                    Else
                        lCoefFranq = mfBuscaCoefFranquiasAdicional(.TextMatrix(.RowSel, ColValLMG), .TextMatrix(.RowSel, ColValFranq))
                    End If
                Else
                    .TextMatrix(.RowSel, ColAlterouFranquia) = ""
                    lCoefFranq = 1
                End If
                .TextMatrix(.RowSel, ColCoefFranq) = lCoefFranq
            End If
                .TextMatrix(.RowSel, ColTextoFranquia) = gM46V999.gfTextoFranquia(IIf(IsNumeric(.TextMatrix(.RowSel, ColValFranq)), .TextMatrix(.RowSel, ColValFranq), 0), .TextMatrix(.RowSel, ColPerFranq), Mid(.TextMatrix(.RowSel, ColTextoFranquia), 1, 1), .TextMatrix(.RowSel, ColCodCobert), val(.TextMatrix(.RowSel, ColPInd)))
                
            Else
                If .ColSel = ColValFranq Or .ColSel = ColValLMG Or .ColSel = ColPInd Then
                
                    Dim lrsFranquia As clsYasRecordSet
                    Dim lCodCobert As Integer
                    Dim lValRiscoCob As Double
                    
                    lCodCobert = .TextMatrix(.RowSel, ColCodCobert)
                    lValRiscoCob = .TextMatrix(.RowSel, ColValLMG)
                
                    Call gM46V999.mfFranqNovoLim(lCodCobert, lValRiscoCob, frmT46V101A.objstcPedLoc.cod_clasf, frmT46V101A.objstcPedLoc.cod_compl_clasf, _
                                                    frmT46V101A.objstcPedLoc.Cod_Constr, Replace(frmT46V101A.txtRisCEP, "-", ""), frmT46V101A.cboTipCobert.ItemData(frmT46V101A.cboTipCobert.ListIndex), _
                                                    IIf(IsNumeric(.TextMatrix(.RowSel, ColPInd)), .TextMatrix(.RowSel, ColPInd), 0), lrsFranquia)
                    If lrsFranquia.RecordCount > 0 Then
                        .TextMatrix(.RowSel, ColCoefFranq) = gM46V999.gfTratarNulo(lrsFranquia("Fator_Ajuste"), tpoNumerico)
                        
                        If .ColSel = ColPInd Then
                            If gM46V999.gfTratarNulo(lrsFranquia("Dias_Franq_Min"), tpoNumerico) > 0 Then
                                .TextMatrix(.RowSel, ColValFranMin) = gM46V999.gfTratarNulo(lrsFranquia("Dias_Franq_Min"), tpoNumerico)
                                .TextMatrix(.RowSel, ColValFranq) = gM46V999.gfTratarNulo(lrsFranquia("Dias_Franq_Min"), tpoNumerico)
                                .TextMatrix(.RowSel, ColMultiploFranquia) = "1x"
                                .TextMatrix(.RowSel, ColDesconto) = "0,00"
                                Call mpAtualizaFranquiaMin(.RowSel)
                                .ColSel = ColPInd
                                .Col = .ColSel
                            End If
                        End If
                        
                    End If
                
                End If
                
            If .TextMatrix(.RowSel, ColCodCobert) = 1691 Then
            
                For i = 1 To .Rows - 1
                    If .TextMatrix(i, ColCodCobert) = 3002 Then
                        .TextMatrix(i, ColLimiteMaximo) = Format(AjustePontualValorMaximo3002("3002", .TextMatrix(i, ColLimiteMaximo), .TextMatrix(i, ColValISmin)), "standard")
                    End If
                Next
            
            End If
            
            Dim Tem_VideObs As String
            
            If gM46V999.gfDisponivel("frmT46V179A") Then
                Tem_VideObs = frmT46V179A.flxCobertura.TextMatrix(frmT46V179A.mpGetFranquiaID(.TextMatrix(.RowSel, ColCodCobert)), 4)
            End If
            
                'lFranqDias = IIf(.TextMatrix(.RowSel, ColTipFranq) = "Dias", .TextMatrix(.RowSel, ColValFranq), 0)
                lFranqDias = IIf(IsNumeric(.TextMatrix(.RowSel, ColPInd)), .TextMatrix(.RowSel, ColPInd), 0)
            
            .TextMatrix(.RowSel, ColTextoFranquia) = gM46V999.gfTextoFranquia(.TextMatrix(.RowSel, ColValFranq), .TextMatrix(.RowSel, ColPerFranq), _
                                                                              Mid(Tem_VideObs, 1, 1), .TextMatrix(.RowSel, ColCodCobert), _
                                                                              lFranqDias, .TextMatrix(.RowSel, ColValRisco), frmT46V101A.objstcPedLoc)
            End If
            
        End Select

        'LMG(LC)
        If .ColSel = ColValLMG Then
            Call mpRecalcularLMGLC
        End If

        Call gpHabSalvar
        .Redraw = True

    End With
End Sub

Private Sub flxSelecionadas_SelChange()
    Dim lI As Integer
    Dim ColAnt As Integer
    Dim RowAnt As Integer
    Dim lCoefFranq As Double
    Dim lCod_Clasf As String
    Dim lCod_Compl_Clasf As String
    Dim lFranqDias As Integer

    'Controle de linha implementado por solicitação da Thais - 23/11/2010
    Dim lRoWSelInicio As Integer

    If Me.Tag = "Desabilitado" Then Exit Sub

    lRoWSelInicio = flxSelecionadas.RowSel

    'Ficha 1000757
    lCod_Clasf = Left(frmT46V101A.txtClassificacao.Tag, 4)
    lCod_Compl_Clasf = Mid(frmT46V101A.txtClassificacao.Tag, 5, 2)

    LstFranquias.Visible = False
    
    With flxSelecionadas

        ColAnt = .Col
        RowAnt = .Row
        .Redraw = False
        TrocandoCor = True

        'Ficha 1000757
        Call VerificaAlteracaoFranquia(lCod_Clasf, lCod_Compl_Clasf)

        If (.ColSel = ColValLMG Or .ColSel = ColValFranq Or .ColSel = ColDescComl Or .ColSel = ColPInd Or .ColSel = ColTaxaFinal Or .ColSel = ColValRisco) And mLinhaFocus < .Rows Then
            If IsNumeric(.TextMatrix(mLinhaFocus, .ColSel)) Then


                'Ficha 1000183
                If .ColSel = ColPInd Then   'thais
                    If .TextMatrix(mLinhaFocus, ColPermitePInd) = "True" And val(.TextMatrix(mLinhaFocus, ColPInd)) = 0 Then
                        If Not gM46O001.PermiteAlterarFranquia Then
                            '.TextMatrix(mLinhaFocus, ColPInd) = 12
                            .TextMatrix(mLinhaFocus, ColPInd) = mfBuscaPerIndenitarioDefault(.TextMatrix(mLinhaFocus, ColCodCobert))
                        End If
                        RowAlterouFranquia = .RowSel
                    End If
                ElseIf .ColSel = ColValFranq Then
                ElseIf .ColSel = ColCoefFranq Then
                    .TextMatrix(mLinhaFocus, .ColSel) = Format(.TextMatrix(mLinhaFocus, .ColSel), "0.00000")
                ElseIf .ColSel = ColValRisco Then
                    .Col = ColValRisco
                    .Row = mLinhaFocus
                    If gM46V999.CDblx(.TextMatrix(mLinhaFocus, ColValRisco)) = 0 Or (.CellBackColor <> vbWhite And .CellBackColor <> &H80000018 And mLinhaFocus <> 1) Then
                        .TextMatrix(mLinhaFocus, ColValRisco) = Format(.TextMatrix(mLinhaFocus, ColValLMG), "standard")
                    Else
                        .TextMatrix(mLinhaFocus, ColValRisco) = Format(.TextMatrix(mLinhaFocus, ColValRisco), "standard")
                    End If
                ElseIf .ColSel = ColTaxaFinal Then
                    If InStr(.TextMatrix(mLinhaFocus, .ColSel), ",") > 0 Then
                        .TextMatrix(mLinhaFocus, .ColSel) = Left(.TextMatrix(mLinhaFocus, .ColSel), InStr(.TextMatrix(mLinhaFocus, .ColSel), ",") - 1) & "," & Mid(.TextMatrix(mLinhaFocus, .ColSel), InStr(.TextMatrix(mLinhaFocus, .ColSel), ",") + 1, 5)
                    End If
                    .TextMatrix(mLinhaFocus, .ColSel) = Format(.TextMatrix(mLinhaFocus, .ColSel), "0.00000")
                Else
                    .TextMatrix(mLinhaFocus, .ColSel) = Format(.TextMatrix(mLinhaFocus, .ColSel), "standard")
                End If
                'If .ColSel = ColValFranq And gM46V999.CDblx(.TextMatrix(gM46V999.CDblx(mLinhaFocus), ColValFranq)) = 0 Then
                '    mpAtualizaFranquiaMin gM46V999.CDblx(mLinhaFocus)
                'End If
                'Ficha 1000183
                If (gM46V999.CDblx(.TextMatrix(mLinhaFocus, ColValFranq)) > 0 And gM46V999.gNovaPropostaLMI) Or _
                   (gM46V999.CDblx(.TextMatrix(mLinhaFocus, ColValFranq)) > 0 And gM46O001.PermiteAlterarFranquia) And mLinhaFocus > 0 And gM46V999.gCodProduto < 1080 Then

                    If gM46V999.CDblx(.TextMatrix(mLinhaFocus, ColValFranq)) <> gM46V999.CDblx(.TextMatrix(mLinhaFocus, ColValFranMin)) Then
                        If .TextMatrix(mLinhaFocus, ColCodCobert) = "1" Then
                            lCoefFranq = mfBuscaCoefFranquiasBasica(.TextMatrix(mLinhaFocus, ColValLMG), .TextMatrix(mLinhaFocus, ColValFranq))
                        Else
                            lCoefFranq = mfBuscaCoefFranquiasAdicional(.TextMatrix(mLinhaFocus, ColValLMG), .TextMatrix(mLinhaFocus, ColValFranq))
                        End If
                    Else
                        lCoefFranq = 1
                    End If
                    .TextMatrix(mLinhaFocus, ColCoefFranq) = lCoefFranq
                End If
            Else
                For lI = 1 To .Rows - 1
                    .TextMatrix(lI, ColValLMG) = Format(.TextMatrix(lI, ColValLMG), "standard")
                    .Col = ColValRisco
                    .Row = lI
                    If gM46V999.CDblx(.TextMatrix(lI, ColValRisco)) = 0 Or (.CellBackColor <> vbWhite And .CellBackColor <> &H80000018 And lI <> 1) Then
                        If .TextMatrix(lI, ColCodCobert) = "1" Then
                            .TextMatrix(lI, ColValRisco) = frmT46V101A.txtValoRisco.Text
                        Else
                            .TextMatrix(lI, ColValRisco) = Format(.TextMatrix(lI, ColValLMG), "standard")
                        End If
                    Else
                        .TextMatrix(lI, ColValRisco) = Format(.TextMatrix(lI, ColValRisco), "standard")
                    End If
                    If .TextMatrix(lI, ColTipFranq) = "Dias" Then
                        .TextMatrix(lI, ColValFranq) = .TextMatrix(lI, ColValFranq)
                    Else
                    .TextMatrix(lI, ColValFranq) = Format(.TextMatrix(lI, ColValFranq), "standard")
                    End If

                    'Ficha 901683
                    'mpAtualizaFranquiaMin lI

                    .TextMatrix(lI, ColDescComl) = Format(.TextMatrix(lI, ColDescComl), "standard")
                    .TextMatrix(lI, ColDescExp) = Format(.TextMatrix(lI, ColDescExp), "standard")
                    If Not IsNumeric(.TextMatrix(lI, ColTaxaFinal)) Then
                        .TextMatrix(lI, ColTaxaFinal) = Format(0, "0.00000")
                    End If
                    If InStr(.TextMatrix(lI, ColTaxaFinal), ",") > 0 Then
                        .TextMatrix(lI, ColTaxaFinal) = Left(.TextMatrix(lI, ColTaxaFinal), InStr(.TextMatrix(lI, ColTaxaFinal), ",") - 1) & "," & Mid(.TextMatrix(lI, ColTaxaFinal), InStr(.TextMatrix(lI, ColTaxaFinal), ",") + 1, 5)
                    End If
                    If .TextMatrix(lI, ColTaxaFinal) > 99 Then
                        .TextMatrix(lI, ColTaxaFinal) = 99
                    End If
                    
                    .TextMatrix(lI, ColTaxaFinal) = Format(.TextMatrix(lI, ColTaxaFinal), "0.00000")
                    'Ficha 1000183  'thais
                    If .TextMatrix(lI, ColPermitePInd) = "True" And val(.TextMatrix(lI, ColPInd)) = 0 Then
                        If Not gM46O001.PermiteAlterarFranquia Then
                            'lFranqDias = IIf(.TextMatrix(lI, ColTipFranq) = "Dias", .TextMatrix(lI, ColValFranq), 0)
                            lFranqDias = IIf(IsNumeric(.TextMatrix(lI, ColPInd)), .TextMatrix(lI, ColPInd), 0)
                
                            '.TextMatrix(lI, ColPInd) = 12
                            .TextMatrix(lI, ColPInd) = mfBuscaPerIndenitarioDefault(.TextMatrix(lI, ColCodCobert))
                            .TextMatrix(lI, ColTextoFranquia) = gM46V999.gfTextoFranquia(.TextMatrix(lI, ColValFranq), .TextMatrix(lI, ColPerFranq), _
                                                                                        Mid(.TextMatrix(lI, ColTextoFranquia), 1, 1), _
                                                                                        .TextMatrix(lI, ColCodCobert), lFranqDias, _
                                                                                        .TextMatrix(lI, ColValRisco), frmT46V101A.objstcPedLoc)
                        End If
                    End If
                    'Ficha 1000183 - sandra
                    If ((gM46V999.CDblx(.TextMatrix(lI, ColValFranq)) > 0 And gM46V999.gNovaPropostaLMI) Or gM46O001.PermiteAlterarFranquia) And gM46V999.gCodProduto < 1080 Then
                        If gM46V999.CDblx(.TextMatrix(lI, ColValFranq)) <> gM46V999.CDblx(.TextMatrix(lI, ColValFranMin)) Then
                            
                            If .TextMatrix(lI, ColValFranq) = Empty Then
                                If .TextMatrix(lI, ColTipFranq) = "Dias" Then
                                    .TextMatrix(lI, ColValFranq) = 0
                                Else
                               .TextMatrix(lI, ColValFranq) = Format("0", "standard")
                            End If
                            End If
                            
                            If .TextMatrix(lI, ColCodCobert) = "1" Then
                                lCoefFranq = mfBuscaCoefFranquiasBasica(.TextMatrix(lI, ColValLMG), .TextMatrix(lI, ColValFranq))
                            Else
                                lCoefFranq = mfBuscaCoefFranquiasAdicional(.TextMatrix(lI, ColValLMG), .TextMatrix(lI, ColValFranq))
                            End If
                        Else
                            lCoefFranq = 1
                        End If
                        .TextMatrix(lI, ColCoefFranq) = lCoefFranq
                    End If
                Next lI
            End If
        End If
        
        If RowIsAlterada <> 0 Then
            mpAtualizaFranquiaMin RowIsAlterada
        End If
        
        If RowAlterouFranquia <> 0 And RowAlterouFranquia < .Rows Then
            lI = RowAlterouFranquia
            
            'lFranqDias = IIf(.TextMatrix(lI, ColTipFranq) = "Dias", .TextMatrix(lI, ColValFranq), 0)
            lFranqDias = IIf(IsNumeric(.TextMatrix(lI, ColPInd)), .TextMatrix(lI, ColPInd), 0)
            .TextMatrix(lI, ColPerFranq) = gM46V999.CDblx(.TextMatrix(lI, ColPerFranq))
                
            .TextMatrix(lI, ColTextoFranquia) = gM46V999.gfTextoFranquia(gM46V999.CDblx(.TextMatrix(lI, ColValFranq)), .TextMatrix(lI, ColPerFranq), _
                                                                        Mid(.TextMatrix(lI, ColTextoFranquia), 1, 1), .TextMatrix(lI, ColCodCobert), _
                                                                        lFranqDias, gM46V999.CDblx(.TextMatrix(lI, ColValRisco)), frmT46V101A.objstcPedLoc)
            
            If gM46V999.gCodProduto >= 1080 Then
                blnRECALCULA_MULTIPLO = False
                
                Dim Aux As String
                Dim i As Integer
                
                i = 1
                Do While i < LstFranquias.ListCount
                    Aux = Trim(LstFranquias.List(i))
                    If gM46V999.CDblx(flxSelecionadas.TextMatrix(lI, ColValFranq)) <= Trim(Mid(Aux, 1, InStr(Aux, "-") - 2)) Then
                        Exit Do
                    End If
                    i = i + 1
                Loop
                
                Aux = IIf(Trim(Mid(Aux, InStr(Aux, "-") + 2)) = Empty, 0, Trim(Mid(Aux, InStr(Aux, "-") + 2)))
                flxSelecionadas.TextMatrix(lI, ColDesconto) = Format(Aux, "standard")
                flxSelecionadas.TextMatrix(lI, ColCoefFranq) = 0
                Aux = IIf(i > 1, "3x", "1x")
                flxSelecionadas.TextMatrix(lI, ColMultiploFranquia) = Aux
                
                'CR93231 - Franquia (permitir que seja selecionado uma franquia menos que o padrão) a partir da 1084
                If gM46V999.CDblx(flxSelecionadas.TextMatrix(lI, ColValFranq)) < gM46V999.CDblx(flxSelecionadas.TextMatrix(lI, ColValFranMin)) And gM46V999.gCodProduto < 1084 Then
                    flxSelecionadas.TextMatrix(lI, ColValFranq) = flxSelecionadas.TextMatrix(lI, ColValFranMin)
                End If
            End If
            
            If flxSelecionadas.TextMatrix(lI, ColValFranMin) <> flxSelecionadas.TextMatrix(lI, ColValFranq) Then
                flxSelecionadas.TextMatrix(lI, ColAlterouFranquia) = "SIM"
            End If
        End If

        RowAlterouFranquia = 0
        RowIsAlterada = 0
        If .RowSel <> lRoWSelInicio Then
            .RowSel = lRoWSelInicio
        End If
        mLinhaFocus = .RowSel
        .ColWidth(ColTabCobert) = 0
        If gM46V999.gCodProduto > gM46V999.TarifaAntiga Then
            If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda And (gM46V999.PerfilAcesso <> e_Perfil_CALC And gM46V999.PerfilAcesso <> e_Perfil_CALC_Gestor) Then
                .ColWidth(ColDescExp) = 1300
                .ColWidth(ColDescComl) = 1300
            Else
                .ColWidth(ColDescExp) = 0
                .ColWidth(ColDescComl) = 0
            End If
        Else
            .ColWidth(ColDescExp) = 1300
            .ColWidth(ColDescComl) = 1300
        End If
        .ColWidth(ColCobObrigat) = 0
        .ColWidth(ColPerLimite) = 0
        .ToolTipText = ""
        If .ColSel = 1 Then
            .ToolTipText = .TextMatrix(.RowSel, ColNomCobert)
        End If
        .Col = ColAnt
        .Row = RowAnt
        
        If Not IsNumeric(.TextMatrix(.RowSel, ColValFranq)) And .RowSel > 1 Then
            .TextMatrix(.RowSel, ColValFranq) = Format(0, "standard")
        End If
        
        .Redraw = True
    End With
End Sub

Private Sub Form_Activate()

    On Error GoTo Desbloquear_Grid

    flxCobertura.Redraw = False
    flxSelecionadas.MousePointer = vbHourglass

    If gM46V999.gCodProduto < 1080 Then
        flxSelecionadas.ColWidth(ColMultiploFranquia) = 0
        flxSelecionadas.ColWidth(ColTipFranq) = 0
        flxSelecionadas.ColWidth(ColDesconto) = 0
    End If

    frmT46V100A.panForm.Caption = Mid$(Me.Name, 4)
    If gM46V999.gChamaSelecionarCoberturas = True And gM46V111.gCod_plano < 2000 Then
        Call mpSelecionarCoberturas(frmT46V101A.objstcPedLoc.ColecaoPedCobert)
        Call frmT46V117A.mpCepAlterado
        gCepAlterado = False
    Else
        gM46V999.gChamaSelecionarCoberturas = False
        If gM46V999.gCodProduto >= 1080 And gM46V111.gCod_plano >= 2000 And frmT46V101A.objstcPedido.getPlano.EhPlanoSegmentado = False Then
            Dim i As Integer
            
            With flxSelecionadas
                .ColWidth(ColMultiploFranquia) = 0
                .ColWidth(ColTipFranq) = 0
                .ColWidth(ColDesconto) = 0
            End With
            
            For i = 1 To flxSelecionadas.Rows - 1
                Call mpAtualizaFranquiaMin(i)
            Next i
        Else
            If gPiDefault And gM46V111.gCod_plano < 2000 Then
                Call ValidaPI
            End If
        End If
    
        If gCepAlterado = True Then
            Call frmT46V117A.mpCepAlterado
            gCepAlterado = False
        End If
    End If

    flxSelecionadas.LeftCol = 0
    If flxSelecionadas.Enabled And flxSelecionadas.Visible Then
        flxSelecionadas.SetFocus
    End If
Finalizar:
    If gM46V999.gDesabilitarForms Or gM46V999.gDesabFormsCotacao Or _
       gM46V111.gCod_Dizer = e_CodDizerCancelamento Or gM46V111.gCod_Dizer = e_CodDizerexclusao Then
        If Me.Tag <> "Desabilitado" Then
            DesabilitaControles Me
            flxSelecionadas.Enabled = True
        End If
    Else
        If Me.Tag = "Desabilitado" Then
            HabilitaControles Me
        End If
    End If
    If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then
        If gM46V999.gDesabFormsTransmi = True Then
            DesabilitaControles Me
            cmdOK.Enabled = True
        End If
    End If
    cmdOK.Enabled = True
    '1200220 - Tarifa 1032 e 1033 Marco 2012 - item 3
    flxCobertura.Col = 1
    flxCobertura.ColSel = 1
    flxCobertura.Sort = 1

Desbloquear_Grid:
    Me.Width = 13500
    Call CentralizarForm
    
    flxCobertura.Redraw = True
    flxSelecionadas.MousePointer = vbDefault
End Sub

Private Sub CentralizarForm()
    With Me
        .Top = (Screen.Height - .Height) / 2
        .Left = (Screen.Width - .Width) / 2
    End With
End Sub

Private Sub Form_Resize()
    
    Const MINWIDTH As Long = 10200
    Const MINHEIGHT As Long = 7050
 
    If Me.Width < MINWIDTH Then
        Me.Width = MINWIDTH
        Exit Sub
    End If
    If Me.Height < MINHEIGHT Then
        Me.Height = MINHEIGHT
        Exit Sub
    End If
    
    If Me.Width < 10200 Or Me.Height < 7050 Then
        If Me.Width < 10200 Then Me.Width = 10200
        If Me.Height < 7050 Then Me.Height = 7050
        
        TabStrip1.Width = 9930
        flxCobertura.Width = 9825
        cmdAdicionar.Left = 3300
        cmdRetirar.Left = 4515
        LBL_ordenaCob.Left = 7530
        CBO_ordenaCob.Left = 8460
        
        TabStrip2.Height = 2610
        TabStrip2.Width = 9930
        flxSelecionadas.Width = 9825
        flxSelecionadas.Height = 2145
        
        cmdFranquia.Top = 5790
        cmdFranquia.Left = 7770
        cmdTexto.Top = 5790
        cmdTexto.Left = 5610
        cmdOK.Top = 5775
        cmdOK.Left = 4418
    Else
        Dim Top As Integer
        Dim Left As Integer
        
        Top = Me.Height - 7050
        Left = Me.Width - 10200
        
        'Parte de cima
        TabStrip1.Width = 9930 + Left
        flxCobertura.Width = 9825 + Left
        cmdAdicionar.Left = 3300 + Left
        cmdRetirar.Left = 4515 + Left
        LBL_ordenaCob.Left = 7530 + Left
        CBO_ordenaCob.Left = 8460 + Left
        
        'Parte de baixo
        TabStrip2.Height = 2610 + Top
        TabStrip2.Width = 9930 + Left
        flxSelecionadas.Height = 2145 + Top
        flxSelecionadas.Width = 9825 + Left
        
        
        cmdTexto.Top = 5790 + Top
        cmdTexto.Left = 5610 + Left
        cmdFranquia.Top = 5790 + Top
        cmdFranquia.Left = 7770 + Left
        cmdOK.Top = 5775 + Top
        cmdOK.Left = 4418 + Left
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    gCepAlterado = False
    mColunaFocus = 0
    mLinhaFocus = 0
    mLinhaFinal = 0
    mLinhaInicial = 0
End Sub


Private Sub LBL_ordenaCob_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If (Button = 2) And (gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda) Then
        Me.Tag = ""
        Me.flxSelecionadas.Enabled = True
    End If
End Sub

Private Sub LstFranquias_Click()
    Dim Aux As String
    Dim ValAnt As String
    Dim Aux2 As String

    If Me.Tag = "Desabilitado" Then Exit Sub

    If PressUpDown Then
        PressUpDown = False
        Exit Sub
    End If
    PressUpDown = False

    If flxSelecionadas.Row > 0 Then

        flxSelecionadas.Col = flxSelecionadas.ColSel
        flxSelecionadas.CellBackColor = vbWhite
        Aux = Trim(LstFranquias.List(LstFranquias.ListIndex))
        Aux2 = IIf(LstFranquias.ListIndex = 1, "1x", "3x")
        'ValAnt = flxSelecionadas.TextMatrix(flxSelecionadas.Row, ColValFranq)
    
        If Aux = "" Then
            If flxSelecionadas.TextMatrix(flxSelecionadas.Row, ColTipFranq) = "Dias" Then
                flxSelecionadas.TextMatrix(flxSelecionadas.Row, ColValFranq) = "0"
            Else
            flxSelecionadas.TextMatrix(flxSelecionadas.Row, ColValFranq) = "0,00"
            End If
            flxSelecionadas.TextMatrix(flxSelecionadas.Row, ColCoefFranq) = 0
        Else
            flxSelecionadas.TextMatrix(flxSelecionadas.Row, ColValFranq) = Trim(Mid(Aux, 1, InStr(Aux, "-") - 2))
            
            If flxSelecionadas.TextMatrix(flxSelecionadas.Row, ColCodGrupoCobert) <> 1 And flxSelecionadas.TextMatrix(flxSelecionadas.Row, ColCodGrupoCobert) <> 5 Then
                flxSelecionadas.TextMatrix(flxSelecionadas.Row, ColMultiploFranquia) = Aux2
                flxSelecionadas.TextMatrix(flxSelecionadas.Row, ColDesconto) = Format(Trim(Mid(Aux, InStr(Aux, "-") + 2)), "standard")
                If val(flxSelecionadas.TextMatrix(flxSelecionadas.Row, ColValFranq)) <> 0 Then
                    flxSelecionadas.TextMatrix(flxSelecionadas.Row, ColCoefFranq) = (100 - val(Trim(Mid(Aux, InStr(Aux, "-") + 2)))) / 100
                Else
                    flxSelecionadas.TextMatrix(flxSelecionadas.Row, ColCoefFranq) = 0
                End If
            
            Else
                flxSelecionadas.TextMatrix(flxSelecionadas.Row, ColMultiploFranquia) = "1x"
                flxSelecionadas.TextMatrix(flxSelecionadas.Row, ColDesconto) = Format("0", "standard")
                flxSelecionadas.TextMatrix(flxSelecionadas.Row, ColCoefFranq) = 0
            End If
            
        End If

        'CR93231 - Franquia (permitir que seja selecionado uma franquia menos que o padrão) a partir da 1084
        If gM46V999.CDblx(flxSelecionadas.TextMatrix(flxSelecionadas.Row, ColValFranq)) < gM46V999.CDblx(flxSelecionadas.TextMatrix(flxSelecionadas.Row, ColValFranMin)) And gM46V999.gCodProduto < 1084 Then
            flxSelecionadas.TextMatrix(flxSelecionadas.Row, ColValFranq) = flxSelecionadas.TextMatrix(flxSelecionadas.Row, ColValFranMin)
        End If
    
        If flxSelecionadas.TextMatrix(flxSelecionadas.Row, ColValFranMin) <> flxSelecionadas.TextMatrix(flxSelecionadas.Row, ColValFranq) Then
            flxSelecionadas.TextMatrix(flxSelecionadas.Row, ColAlterouFranquia) = "SIM"
        End If
        
        LstFranquias.Visible = False
        flxSelecionadas_SelChange
        flxSelecionadas.Col = flxSelecionadas.Col - 1
        flxSelecionadas.SetFocus
        With flxSelecionadas
            Dim lFranqDias As Integer
            'lFranqDias = IIf(.TextMatrix(.RowSel, ColTipFranq) = "Dias", .TextMatrix(.RowSel, ColValFranq), 0)
            lFranqDias = IIf(IsNumeric(.TextMatrix(.RowSel, ColPInd)), .TextMatrix(.RowSel, ColPInd), 0)
                
            .TextMatrix(.RowSel, ColTextoFranquia) = gM46V999.gfTextoFranquia(.TextMatrix(.RowSel, ColValFranq), .TextMatrix(.RowSel, ColPerFranq), _
                                                                                Mid(.TextMatrix(.RowSel, ColTextoFranquia), 1, 1), _
                                                                                .TextMatrix(.RowSel, ColCodCobert), lFranqDias, _
                                                                                gM46V999.CDblx(.TextMatrix(.RowSel, ColValRisco)), frmT46V101A.objstcPedLoc)
        End With
    
    End If
End Sub

Private Sub LstFranquias_DblClick()
    LstFranquias_Click
End Sub

Private Sub LstFranquias_KeyDown(KeyCode As Integer, Shift As Integer)

    If Me.Tag = "Desabilitado" Then Exit Sub

    If KeyCode = 13 Then
        LstFranquias_Click
    End If
    If KeyCode = 27 Then
        LstFranquias.Visible = False
        flxSelecionadas.SetFocus
    End If
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        PressUpDown = True
        'Ficha 1000304
    ElseIf gM46O001.PermiteAlterarFranquia Then
        LstFranquias.Visible = False
        flxSelecionadas.SetFocus

    End If

End Sub

Private Sub LstFranquias_LostFocus()
    If LstFranquias.Visible Then
        LstFranquias.Visible = False
    End If
End Sub

Private Sub lstMultiploFranquias_Click()

End Sub

Private Sub txtDescComercial_Change()
    Call gpLimPremios
End Sub

Private Sub txtDescComercial_GotFocus()

    If Me.Tag = "Desabilitado" Then Exit Sub
    With txtDescComercial
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtDescComercial_KeyDown(KeyCode As Integer, Shift As Integer)

    If Me.Tag = "Desabilitado" Then Exit Sub

    If KeyCode = 13 Then
        txtDescComercial_LostFocus
        Call cmdMover_Click
    End If

End Sub

Private Sub txtDescComercial_LostFocus()

    If Me.Tag = "Desabilitado" Then Exit Sub

    With txtDescComercial
        If gM46V999.gfPreenchido(.Text) Then
            If IsNumeric(.Text) Then
                .Text = Format(.Text, "standard")
            Else
                .Text = Format(0, "standard")
            End If
        Else
            .Text = Format(0, "standard")
        End If
        If val(.Text) >= 100 Then
            .Text = "99,99"
        End If
    End With
End Sub

Private Sub Form_Load()
    Dim lMensagem As String
    '1. Posiciona form.
    Call gM46V999.gpCenForm(Me)
    
    blnRECALCULA_MULTIPLO = False
    
    'Acesso Syas Privilégios
    'gM46O001.PermiteAlterarFranquia = (gM46V999.gCodProduto >= 1076)
    
    With flxCobertura
        .Rows = 1
        .Cols = 9
        .TextMatrix(0, 0) = "Código"
        If gM46V999.gCod_Ramo = e_CodRamo_Roubos Then
            .TextMatrix(0, 1) = "Bem coberto"
            TabStrip1.Tabs(1).Caption = "Bens cobertos"
            TabStrip2.Tabs(1).Caption = "Bens contratados"

        Else
            .TextMatrix(0, 1) = "Cobertura"
            TabStrip1.Tabs(1).Caption = "Coberturas"
            TabStrip2.Tabs(1).Caption = "Coberturas contratadas"
        End If
        .TextMatrix(0, 2) = "Limite máximo"
        .TextMatrix(0, 3) = "TAB_COBERT"
        .TextMatrix(0, 4) = "PER_FRANQ"
        .TextMatrix(0, 5) = "VAL_FRANQ_MIN"
        .TextMatrix(0, 6) = "Cobertura obrigatória"
        .TextMatrix(0, 7) = "Cobertura obrigatória"
        .Row = 0
        .ColWidth(0) = 960
        .ColWidth(1) = 6960
        .ColWidth(2) = 1515
        .ColWidth(3) = 0
        .ColWidth(4) = 0
        .ColWidth(5) = 0
        .ColWidth(6) = 0
        .ColWidth(7) = 0
        .ColWidth(8) = 0    'ColCodGrupoCobert
    End With

    With flxSelecionadas
        .Rows = 1
        .Cols = 29
        .TextMatrix(0, ColCodCobert) = "Cód"
        If gM46V999.gCod_Ramo = e_CodRamo_Roubos Then
            .TextMatrix(0, ColNomCobert) = "Bem coberto"
        Else
            .TextMatrix(0, ColNomCobert) = "Cobertura"
        End If
        .TextMatrix(0, ColValLMG) = "Lim.Max.Indenização"
        .TextMatrix(0, ColLimiteMaximo) = "Limite máximo"
        .TextMatrix(0, ColValRisco) = "Valor Risco"
        .TextMatrix(0, ColTabCobert) = "TAB_COBERT"
        .TextMatrix(0, ColPerFranq) = "%Franq"
        .TextMatrix(0, ColTipFranq) = "#"
        .TextMatrix(0, ColValFranq) = "Franquia Min"
        .TextMatrix(0, ColDescComl) = "Des. Comercial"
        .TextMatrix(0, ColDescExp) = "Des. Experiência"
        .TextMatrix(0, ColCobObrigat) = "cobertura obrigatória"
        .TextMatrix(0, ColPerLimite) = "Per_Limite"
        .TextMatrix(0, ColPermitePInd) = "Permite Período Indenitario"
        .TextMatrix(0, ColPInd) = "PI"
        .TextMatrix(0, ColCoefFranq) = "Coef.Franq"
        .TextMatrix(0, ColTaxaFinal) = "Taxa Final"
        .TextMatrix(0, ColTaxaAlterada) = "TaxaAlterada"
        .TextMatrix(0, ColValFranMin) = "ValFranMin"
        .TextMatrix(0, ColMultiploFranquia) = "Múltiplo"
        .TextMatrix(0, ColDesconto) = "%Desc"
        .TextMatrix(0, ColCobertPerLimite) = "CobertPerLimite"
        .TextMatrix(0, ColValISAte) = "IS Até"
        .TextMatrix(0, ColAlterouFranquia) = "FranAlterada"
        .TextMatrix(0, ColDescFranquia) = "NomecobFranquia"
        .TextMatrix(0, ColTextoFranquia) = "Texto da Franquia"
        .TextMatrix(0, ColValISmin) = "IS Min"
        .ColWidth(ColCodGrupoCobert) = 0

        If gM46V999.gCodProduto < 1080 Then
            .ColWidth(ColMultiploFranquia) = 0
            .ColWidth(ColTipFranq) = 0
            .ColWidth(ColDesconto) = 0
        End If
        
        .Row = 0
        .ColWidth(ColTipFranq) = 420
        .ColWidth(ColCodCobert) = 495
        .ColWidth(ColNomCobert) = 2640
        .ColWidth(ColValLMG) = 1605
        .ColWidth(ColMultiploFranquia) = 720
        .ColWidth(ColDesconto) = 690

        If gM46V999.gCod_Ramo = e_CodRamo_Roubos Then
            .ColWidth(ColLimiteMaximo) = 0
        Else
            .ColWidth(ColLimiteMaximo) = 1400
        End If
        '        Coluna inibida conforme documentação "F012 - Declaração de Escopo - Revisão de Formulários (Valor em Risco)".
        '        If gM46V999.gCodProduto > gM46V999.TarifaAntiga And gM46V999.gCod_Ramo <> e_CodRamo_Roubos Then
        '            .ColWidth(ColValRisco) = 1400
        '        Else
        .ColWidth(ColValRisco) = 0
        '        End If
        .ColWidth(ColTabCobert) = 0
        .ColWidth(ColPerFranq) = 630
        .ColWidth(ColValFranq) = 1200
        If gM46V999.gCodProduto > gM46V999.TarifaAntiga Then
            If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda And (gM46V999.PerfilAcesso <> e_Perfil_CALC And gM46V999.PerfilAcesso <> e_Perfil_CALC_Gestor) Then
                .ColWidth(ColDescExp) = 1300
                .ColWidth(ColDescComl) = 1300
            Else
                .ColWidth(ColDescExp) = 0
                .ColWidth(ColDescComl) = 0
            End If
        Else
            .ColWidth(ColDescExp) = 1300
            .ColWidth(ColDescComl) = 1300
        End If
        .ColWidth(ColCobObrigat) = 0
        .ColWidth(ColPerLimite) = 0
        .ColWidth(ColPermitePInd) = 0
        .ColWidth(ColPInd) = 300
        .ColWidth(ColCoefFranq) = 0
        If (gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda And gM46O001.PermiteAlterarTaxaFinal) Then
            .ColWidth(ColTaxaFinal) = 1000
            .ColWidth(ColAlterouFranquia) = 1000
        Else
            .ColWidth(ColTaxaFinal) = 0
            .ColWidth(ColAlterouFranquia) = 0
        End If
        .ColWidth(ColTaxaAlterada) = 0
        .ColWidth(ColValFranMin) = 0
        .ColWidth(ColCobertPerLimite) = 0
        .ColWidth(ColValISAte) = 0
        .ColWidth(ColValISmin) = 0
        .ColWidth(ColAjusteEndosso) = 0
        .ColWidth(ColDescFranquia) = 0
        .ColWidth(ColTextoFranquia) = 0
    End With

    'Limpando as variáveis
    Set mColecaoAdicional = Nothing
    mUltProdutoA = 0
    mUltPlanoA = 0
    'CR23794 - Syas Kit - Erro PI
    gPiDefault = False

    Set mColecaoBasica = Nothing
    mUltPlanoB = 0
    mUltProdutoB = 0
    mUltValorLMGB = 0

    Call mpCarregarGridCoberturas
    
    If gM46V999.gCod_Ramo <> e_CodRamo_Empresarial Then
        fraCobertura.Width = 4575
    End If
    If gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Or gM46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda Then
        If frmT46V101A.objstcPedLoc.Tip_Emissao = e_TipEmissao_RenYasuda Then
            cmdBonus.Enabled = False
        Else
            cmdBonus.Enabled = True
        End If
    End If

    cmdFranquia.Visible = False
    If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then
        cmdFranquia.Visible = True
    End If

End Sub

Public Sub mpIncluirCoberturasCapitalGlobal()
    Dim lSelect As String
    Dim Aux As Double
    Dim lbdTAB_Cobert As ADODB.Connection    'Arquivo: P00Tarifa.
    Dim lMensagem As String    'Mensagem.
    Dim lrsTab_Cobert As clsYasRecordSet    'Registro : TAB_COBERT.

    Call mpExcluirCoberturasCapitalGlobal(True)

    If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTAB_Cobert, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If

    lSelect = " Select distinct C.DSC_Cobert, C.Cod_Cobert, C.Per_Lim_Max_Cobert, C.TAB_COBERT, C.PER_FRANQ, C.VAL_FRANQ_MIN, C.COD_LIM_COBERT, C.NUM_ORDEM_COBERT, C.COD_GRUPO_COBERT, V.Tip_Plano_Vida from Tbm_Plano_Vida AS V"
    lSelect = lSelect + " INNER JOIN TAB_COBERT AS C ON C.COD_COBERT = V.Cod_Cobert and C.COD_RAMO = V.Cod_Ramo and C.COD_TARIF_INI = V.Cod_Tarif and C.COD_TARIF_TER = V.Cod_Tarif"
    lSelect = lSelect & " Where C.cod_Ramo = " & gM46V999.gCod_Ramo
    If val(Right(frmT46V101A.txtRisRegiao.Tag, 4)) <> 0 And gM46V999.gCodProduto <= gM46V999.TarifaAntiga Then
        lSelect = lSelect & " and C.tab_cobert = " & val(Right(frmT46V101A.txtRisRegiao.Tag, 4))
    Else
        lSelect = lSelect & " and C.tab_cobert = " & frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex)
    End If
    lSelect = lSelect & " and COD_TARIF_INI <= " & gM46V999.gCodProduto
    lSelect = lSelect & " and COD_TARIF_TER >= " & gM46V999.gCodProduto
    lSelect = lSelect & " and v.Cod_Plano_Vida = " & frmT46V181A.cboCobertura.ListIndex

    If gM46V999.gfObtRegistro(lbdTAB_Cobert, lSelect, lrsTab_Cobert, lMensagem) = False Then
        Call gM46V999.gpFecha2(lbdTAB_Cobert)
        Call gM46V999.gpGraLog(2, lMensagem)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Aux = 0
    Do While Not lrsTab_Cobert.EOF
        flxSelecionadas.Rows = flxSelecionadas.Rows + 1
        flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColCodCobert) = lrsTab_Cobert("Cod_Cobert")
        flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColNomCobert) = Trim(lrsTab_Cobert("DSC_Cobert"))

        Aux = frmT46V181A.txtQtdeVidas.Text * frmT46V181A.txtLMGPorVida.Text
        '        If frmT46V181A.optInclusaoDeConjuge(0).Value = True And lrsTab_Cobert("Tip_Plano_Vida") = 1 Then
        '            Aux = Aux + (frmT46V181A.txtQtdeVidas.Text * frmT46V181A.txtLMGPorVida.Text * 0.5)
        '        End If
        flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColValLMG) = Format(Aux, "standard")

        flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColLimiteMaximo) = Format(Aux, "standard")
        flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColTabCobert) = lrsTab_Cobert("TAB_COBERT")
        flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColPerFranq) = lrsTab_Cobert("PER_FRANQ")
        flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColTipFranq) = "R$"
        'ColValFranq
        flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColValFranq) = Format(lrsTab_Cobert("VAL_FRANQ_MIN"), "standard")
        '        flxSelecionadas.Col = ColValFranq
        '        flxSelecionadas.Row = flxSelecionadas.Rows - 1
        '        flxSelecionadas.CellBackColor = &H80000018
        'ColDescComl
        flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColDescComl) = "0,00"
        '        flxSelecionadas.Col = ColDescComl
        '        flxSelecionadas.Row = flxSelecionadas.Rows - 1
        '        flxSelecionadas.CellBackColor = &H80000018
        'ColDescExp
        flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColDescExp) = "0,00"
        ''        flxSelecionadas.Col = ColDescExp
        ''        flxSelecionadas.Row = flxSelecionadas.Rows - 1
        ''        flxSelecionadas.CellBackColor = &H80000018

        flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColCobObrigat) = 0
        flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColPerLimite) = lrsTab_Cobert("Per_Lim_Max_Cobert")
        flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColPermitePInd) = False
        flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColCobertPerLimite) = lrsTab_Cobert("COD_LIM_COBERT")
        flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColValISAte) = mfRetornaLMG(flxSelecionadas.Rows - 1)
        flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColValISmin) = mfRetornaLMGMin(flxSelecionadas.Rows - 1)
        flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColCodGrupoCobert) = lrsTab_Cobert("COD_GRUPO_COBERT")
        flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColDesconto) = "0,00"
        flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColMultiploFranquia) = "1x"
                
        Call mpCalcularLimite(flxSelecionadas.Rows - 1)
        flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColValRisco) = flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColLimiteMaximo)
        lrsTab_Cobert.MoveNext
    Loop

    '4 - AUXILIO FUNERAL
    If frmT46V181A.optAuxilioFuneral(0).Value = True Then

        lSelect = " Select distinct C.DSC_Cobert, C.Cod_Cobert, C.Per_Lim_Max_Cobert, C.TAB_COBERT, C.PER_FRANQ, C.VAL_FRANQ_MIN, C.COD_LIM_COBERT, C.NUM_ORDEM_COBERT, C.COD_GRUPO_COBERT, V.Tip_Plano_Vida from Tbm_Plano_Vida AS V"
        lSelect = lSelect + " INNER JOIN TAB_COBERT AS C ON C.COD_COBERT = V.Cod_Cobert and C.COD_RAMO = V.Cod_Ramo and C.COD_TARIF_INI = V.Cod_Tarif and C.COD_TARIF_TER = V.Cod_Tarif"
        lSelect = lSelect & " Where C.cod_Ramo = " & gM46V999.gCod_Ramo
        If val(Right(frmT46V101A.txtRisRegiao.Tag, 4)) <> 0 And gM46V999.gCodProduto <= gM46V999.TarifaAntiga Then
            lSelect = lSelect & " and C.tab_cobert = " & val(Right(frmT46V101A.txtRisRegiao.Tag, 4))
        Else
            lSelect = lSelect & " and C.tab_cobert = " & frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex)
        End If
        lSelect = lSelect & " and COD_TARIF_INI <= " & gM46V999.gCodProduto
        lSelect = lSelect & " and COD_TARIF_TER >= " & gM46V999.gCodProduto
        lSelect = lSelect & " and v.Cod_Plano_Vida = 4"

        If gM46V999.gfObtRegistro(lbdTAB_Cobert, lSelect, lrsTab_Cobert, lMensagem) = False Then
            Call gM46V999.gpFecha2(lbdTAB_Cobert)
            Call gM46V999.gpGraLog(2, lMensagem)
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        Aux = 0
        If Not lrsTab_Cobert.EOF Then
            flxSelecionadas.Rows = flxSelecionadas.Rows + 1
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColCodCobert) = lrsTab_Cobert("Cod_Cobert")
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColNomCobert) = Trim(lrsTab_Cobert("DSC_Cobert"))

            Aux = frmT46V181A.txtQtdeVidas.Text * frmT46V181A.txtLMGFuneralPorVida.Text
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColValLMG) = Format(Aux, "standard")

            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColLimiteMaximo) = Format(Aux, "standard")
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColTabCobert) = lrsTab_Cobert("TAB_COBERT")
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColPerFranq) = lrsTab_Cobert("PER_FRANQ")
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColTipFranq) = "R$"
            'ColValFranq
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColValFranq) = Format(lrsTab_Cobert("VAL_FRANQ_MIN"), "standard")
            flxSelecionadas.Col = ColValFranq
            flxSelecionadas.Row = flxSelecionadas.Rows - 1
            flxSelecionadas.CellBackColor = &H80000018
            'ColDescComl
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColDescComl) = "0,00"
            flxSelecionadas.Col = ColDescComl
            flxSelecionadas.Row = flxSelecionadas.Rows - 1
            flxSelecionadas.CellBackColor = &H80000018
            'ColDescExp
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColDescExp) = "0,00"
            flxSelecionadas.Col = ColDescExp
            flxSelecionadas.Row = flxSelecionadas.Rows - 1
            flxSelecionadas.CellBackColor = &H80000018

            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColCobObrigat) = 0
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColPerLimite) = lrsTab_Cobert("Per_Lim_Max_Cobert")
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColPermitePInd) = False
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColCobertPerLimite) = lrsTab_Cobert("COD_LIM_COBERT")
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColValISAte) = mfRetornaLMG(flxSelecionadas.Rows - 1)
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColValISmin) = mfRetornaLMGMin(flxSelecionadas.Rows - 1)
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColCodGrupoCobert) = lrsTab_Cobert("COD_GRUPO_COBERT")
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColDesconto) = "0,00"
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColMultiploFranquia) = "1x"
            Call mpCalcularLimite(flxSelecionadas.Rows - 1)
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColValRisco) = flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColLimiteMaximo)
        End If
    End If

    '4 - MORTE (CONJUGE)
    If frmT46V181A.optInclusaoDeConjuge(0) = True Then

        lSelect = " Select distinct C.DSC_Cobert, C.Cod_Cobert, C.Per_Lim_Max_Cobert, C.TAB_COBERT, C.PER_FRANQ, C.VAL_FRANQ_MIN, C.COD_LIM_COBERT, C.NUM_ORDEM_COBERT, C.COD_GRUPO_COBERT, V.Tip_Plano_Vida from Tbm_Plano_Vida AS V"
        lSelect = lSelect + " INNER JOIN TAB_COBERT AS C ON C.COD_COBERT = V.Cod_Cobert and C.COD_RAMO = V.Cod_Ramo and C.COD_TARIF_INI = V.Cod_Tarif and C.COD_TARIF_TER = V.Cod_Tarif"
        lSelect = lSelect & " Where C.cod_Ramo = " & gM46V999.gCod_Ramo
        If val(Right(frmT46V101A.txtRisRegiao.Tag, 4)) <> 0 And gM46V999.gCodProduto <= gM46V999.TarifaAntiga Then
            lSelect = lSelect & " and C.tab_cobert = " & val(Right(frmT46V101A.txtRisRegiao.Tag, 4))
        Else
            lSelect = lSelect & " and C.tab_cobert = " & frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex)
        End If
        lSelect = lSelect & " and COD_TARIF_INI <= " & gM46V999.gCodProduto
        lSelect = lSelect & " and COD_TARIF_TER >= " & gM46V999.gCodProduto
        lSelect = lSelect & " and v.Cod_Plano_Vida = 6"

        If gM46V999.gfObtRegistro(lbdTAB_Cobert, lSelect, lrsTab_Cobert, lMensagem) = False Then
            Call gM46V999.gpFecha2(lbdTAB_Cobert)
            Call gM46V999.gpGraLog(2, lMensagem)
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        Aux = 0
        If Not lrsTab_Cobert.EOF Then
            flxSelecionadas.Rows = flxSelecionadas.Rows + 1
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColCodCobert) = lrsTab_Cobert("Cod_Cobert")
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColNomCobert) = Trim(lrsTab_Cobert("DSC_Cobert"))

            Aux = frmT46V181A.txtQtdeVidas.Text * (frmT46V181A.txtLMGPorVida.Text / 2)
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColValLMG) = Format(Aux, "standard")

            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColLimiteMaximo) = Format(Aux, "standard")
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColTabCobert) = lrsTab_Cobert("TAB_COBERT")
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColPerFranq) = lrsTab_Cobert("PER_FRANQ")
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColTipFranq) = "R$"
            'ColValFranq
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColValFranq) = Format(lrsTab_Cobert("VAL_FRANQ_MIN"), "standard")
            flxSelecionadas.Col = ColValFranq
            flxSelecionadas.Row = flxSelecionadas.Rows - 1
            flxSelecionadas.CellBackColor = &H80000018
            'ColDescComl
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColDescComl) = "0,00"
            flxSelecionadas.Col = ColDescComl
            flxSelecionadas.Row = flxSelecionadas.Rows - 1
            flxSelecionadas.CellBackColor = &H80000018
            'ColDescExp
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColDescExp) = "0,00"
            flxSelecionadas.Col = ColDescExp
            flxSelecionadas.Row = flxSelecionadas.Rows - 1
            flxSelecionadas.CellBackColor = &H80000018

            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColCobObrigat) = 0
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColPerLimite) = lrsTab_Cobert("Per_Lim_Max_Cobert")
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColPermitePInd) = False
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColCobertPerLimite) = lrsTab_Cobert("COD_LIM_COBERT")
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColValISAte) = mfRetornaLMG(flxSelecionadas.Rows - 1)
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColValISmin) = mfRetornaLMGMin(flxSelecionadas.Rows - 1)
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColCodGrupoCobert) = lrsTab_Cobert("COD_GRUPO_COBERT")
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColDesconto) = "0,00"
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColMultiploFranquia) = "1x"
            Call mpCalcularLimite(flxSelecionadas.Rows - 1)
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColValRisco) = flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColLimiteMaximo)
        End If
    End If

    '5 - AUXILIO CESTA BASICA
    If frmT46V181A.cboCobertura.ListIndex = 3 Then

        lSelect = " Select distinct C.DSC_Cobert, C.Cod_Cobert, C.Per_Lim_Max_Cobert, C.TAB_COBERT, C.PER_FRANQ, C.VAL_FRANQ_MIN, C.COD_LIM_COBERT, C.NUM_ORDEM_COBERT, C.COD_GRUPO_COBERT, V.Tip_Plano_Vida from Tbm_Plano_Vida AS V"
        lSelect = lSelect + " INNER JOIN TAB_COBERT AS C ON C.COD_COBERT = V.Cod_Cobert and C.COD_RAMO = V.Cod_Ramo and C.COD_TARIF_INI = V.Cod_Tarif and C.COD_TARIF_TER = V.Cod_Tarif"
        lSelect = lSelect & " Where C.cod_Ramo = " & gM46V999.gCod_Ramo
        If val(Right(frmT46V101A.txtRisRegiao.Tag, 4)) <> 0 And gM46V999.gCodProduto <= gM46V999.TarifaAntiga Then
            lSelect = lSelect & " and C.tab_cobert = " & val(Right(frmT46V101A.txtRisRegiao.Tag, 4))
        Else
            lSelect = lSelect & " and C.tab_cobert = " & frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex)
        End If
        lSelect = lSelect & " and COD_TARIF_INI <= " & gM46V999.gCodProduto
        lSelect = lSelect & " and COD_TARIF_TER >= " & gM46V999.gCodProduto
        lSelect = lSelect & " and v.Cod_Plano_Vida = 5"

        If gM46V999.gfObtRegistro(lbdTAB_Cobert, lSelect, lrsTab_Cobert, lMensagem) = False Then
            Call gM46V999.gpFecha2(lbdTAB_Cobert)
            Call gM46V999.gpGraLog(2, lMensagem)
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        Aux = 0
        If Not lrsTab_Cobert.EOF Then
            flxSelecionadas.Rows = flxSelecionadas.Rows + 1
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColCodCobert) = lrsTab_Cobert("Cod_Cobert")
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColNomCobert) = Trim(lrsTab_Cobert("DSC_Cobert"))

            Aux = frmT46V181A.txtQtdeVidas.Text * frmT46V181A.txtAuxCestaBasPorVida.Text
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColValLMG) = Format(Aux, "standard")

            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColLimiteMaximo) = Format(Aux, "standard")
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColTabCobert) = lrsTab_Cobert("TAB_COBERT")
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColPerFranq) = lrsTab_Cobert("PER_FRANQ")
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColTipFranq) = "R$"
            'ColValFranq
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColValFranq) = Format(lrsTab_Cobert("VAL_FRANQ_MIN"), "standard")
            flxSelecionadas.Col = ColValFranq
            flxSelecionadas.Row = flxSelecionadas.Rows - 1
            flxSelecionadas.CellBackColor = &H80000018
            'ColDescComl
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColDescComl) = "0,00"
            flxSelecionadas.Col = ColDescComl
            flxSelecionadas.Row = flxSelecionadas.Rows - 1
            flxSelecionadas.CellBackColor = &H80000018
            'ColDescExp
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColDescExp) = "0,00"
            flxSelecionadas.Col = ColDescExp
            flxSelecionadas.Row = flxSelecionadas.Rows - 1
            flxSelecionadas.CellBackColor = &H80000018

            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColCobObrigat) = 0
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColPerLimite) = lrsTab_Cobert("Per_Lim_Max_Cobert")
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColPermitePInd) = False
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColCobertPerLimite) = lrsTab_Cobert("COD_LIM_COBERT")
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColValISAte) = mfRetornaLMG(flxSelecionadas.Rows - 1)
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColValISmin) = mfRetornaLMGMin(flxSelecionadas.Rows - 1)
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColCodGrupoCobert) = lrsTab_Cobert("COD_GRUPO_COBERT")
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColDesconto) = "0,00"
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColMultiploFranquia) = "1x"
            Call mpCalcularLimite(flxSelecionadas.Rows - 1)
            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColValRisco) = flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColLimiteMaximo)
        End If
    End If

End Sub

Public Sub mpExcluirCoberturasCapitalGlobal(Optional ByVal lblForcaExclusao As Boolean = False)
    Dim lI As Integer

    lI = 1
    Do While lI < flxSelecionadas.Rows
        If flxSelecionadas.TextMatrix(lI, ColCodGrupoCobert) = "4" Then
            If gM46V111.gCod_Dizer <> 0 And Not lblForcaExclusao Then
                'Endosso não pode excluir a cobertura
                flxSelecionadas.TextMatrix(lI, ColValLMG) = 0
                lI = lI + 1
            Else
                flxSelecionadas.RemoveItem lI
            End If
        Else
            lI = lI + 1
        End If
    Loop
End Sub

Public Sub mpSepararCoberturas()
    Dim lLinha As Integer
    Dim lPredioConteudoVerbasSeparadas As Boolean

    If (frmT46V101A.cboTipCobert.ListIndex = e_TipCobert_PredioeConteudo) And (frmT46V101A.optVerbasSeparadas(0).Value = True) Then
        lPredioConteudoVerbasSeparadas = True
    Else
        lPredioConteudoVerbasSeparadas = False
    End If

    If lPredioConteudoVerbasSeparadas = True Then

        If flxSelecionadas.Rows < 2 Then
            Call mpCarregarGridCoberturas
        Else
            Call mpCarregarColecaoCobert
        End If

        lLinha = 1
        'Criar cópia da cobertura a ser dividida
        Dim i As Integer
        Dim strCopia As String
        Dim blnExisteCoberturaVerbaSeparada As Boolean
        
        blnExisteCoberturaVerbaSeparada = False
        For i = 1 To flxSelecionadas.Rows - 1
            If flxSelecionadas.TextMatrix(i, ColCodGrupoCobert) = 5 Then
                blnExisteCoberturaVerbaSeparada = True
                Exit For
            End If
        Next
        
        If Not blnExisteCoberturaVerbaSeparada Then
            strCopia = Empty
            For i = 0 To flxSelecionadas.Cols - 1
                strCopia = strCopia + flxSelecionadas.TextMatrix(lLinha, i)
                If i < flxSelecionadas.Cols Then
                    strCopia = strCopia & vbTab
                End If
            Next i
    
            flxSelecionadas.AddItem strCopia, 2
        End If
    
        'Ajusta valor e descrição da parcela Predio
        flxSelecionadas.TextMatrix(lLinha, ColNomCobert) = flxSelecionadas.TextMatrix(lLinha, ColNomCobert) & DESC_COBERT_PREDIO

        If val(flxSelecionadas.TextMatrix(lLinha, ColValLMG)) = 0 Then
            flxSelecionadas.TextMatrix(lLinha, ColValLMG) = Format(frmT46V101A.txtISBas.Text, "standard")
        End If
        flxSelecionadas.TextMatrix(lLinha, ColValLMG) = Format(flxSelecionadas.TextMatrix(lLinha, ColValLMG) / 2, "standard")

        'Ajusta valor e descrição da parcela Conteudo
        lLinha = 2
        Dim cobertura As stcA46V101B
        Set cobertura = New stcA46V101B

        For Each cobertura In ColecaoCobert
            If cobertura.CodGrupoCobert = 5 Then
                flxSelecionadas.TextMatrix(lLinha, ColCodCobert) = cobertura.codigo
                flxSelecionadas.TextMatrix(lLinha, ColNomCobert) = cobertura.Descricao & DESC_COBERT_CONTEUDO
                flxSelecionadas.TextMatrix(lLinha, ColTextoFranquia) = Empty
                Exit For
            End If
        Next cobertura

        If val(flxSelecionadas.TextMatrix(lLinha, ColValLMG)) = 0 Then
            flxSelecionadas.TextMatrix(lLinha, ColValLMG) = Format(frmT46V101A.txtISBas.Text, "standard")
        End If

        flxSelecionadas.TextMatrix(lLinha, ColValLMG) = Format(flxSelecionadas.TextMatrix(lLinha, ColValLMG) / 2, "standard")
        flxSelecionadas.TextMatrix(lLinha, ColCodGrupoCobert) = 5
        
        Call mpVerificarLimiteDependente
        Call mpAtualizaFranquiaMin(lLinha)

    Else
        If flxSelecionadas.Rows < 3 Then Exit Sub

        'Procura cobertura de verba separada
        For lLinha = 1 To flxSelecionadas.Rows - 1
            If flxSelecionadas.TextMatrix(lLinha, ColCodGrupoCobert) = 5 Then
                If gM46V111.gCod_Dizer <> 0 Then
                    'Endosso não pode excluir a cobertura
                    flxSelecionadas.TextMatrix(lLinha, ColValLMG) = 0
                Else
                    flxSelecionadas.RemoveItem lLinha
                End If
    
                'Ajusta valor e descrição da parcela Predio
                lLinha = 1
                flxSelecionadas.TextMatrix(lLinha, ColNomCobert) = ColecaoCobert.Item(1).Descricao
                flxSelecionadas.TextMatrix(lLinha, ColValLMG) = Format(frmT46V101A.txtISBas.Text, "standard")
                Exit For
            End If
        Next
    End If
End Sub

Public Sub mpSelecionarCoberturas(pColecaoPedCobert As Collection, Optional ByVal Calculando As Boolean, Optional ByVal pCarregando As Boolean, Optional ByVal pAtualizarFranqMin As Boolean = False, Optional ByVal pCargaNormal As Boolean = True)
    Dim lobjstcCobert As stcA46V704B    'Objeto da classe de estrutura stcA46V704B (TAB_PED_COBERT).
    Dim lobjstc As stcA46V101B          'Objeto da classe de estrutura stcA46V101A.
    Dim lblnAchou As Boolean            'Indica se achou a cobertura ou não.
    Dim lValLMG As Double               'Valor LMG.
    Dim lDescComercial As Double        'Desconto comercial.
    Dim lDescExperiencia As Double      'Desconto de experiência.
    Dim lPercent As Double
    'Dim lPercentMin                        As Double
    Dim lFranqMinimo As Double
    Dim lMinimoDias As Integer
    Dim lValFatorFraqMin As Double
    Dim lPerMin As Double
    Dim lFranqSelec As Double
    Dim lValFatorFraqSel As Double
    'Dim lValFatorFraqMin                   As Double
    Dim lValRiscoCob As Double
    Dim lNum_Meses_PI As Integer
    Dim lTaxaFinal As Double
    Dim TaxaAlterada As String
    Dim lAlterouFranquia As String
    Dim lAjusteEndosso As String
    Dim FranquiaAlterada As String
    Dim lTextoFranquia As String
    Dim pTextoFranquia As String
    Dim lObrig As Boolean
    Dim lCargaAux As Boolean
    Dim lValorLMGTotal As Double
    Dim lCod_Clasf As String
    Dim lCod_Compl_Clasf As String
    Dim Dummy As Double
    Dim Dummy1 As Integer
    Dim dymmy1 As String
    Dim dValISFixa As Double
    Dim dValISLimite As Double
    Dim lMensagem As String
    Dim codCobertAnt As Integer
    Dim lComplementoDescricaoCobertura As String
    Dim lPredioConteudoVerbasSeparadas As Boolean
    Dim lCodGrupoCobertura As String
    Dim pMuliploFraq As Integer
    Dim pDescMultFraq As Double
    Dim lMultiplo As Integer
    Dim lDescMultiplo As Double
    Dim lAtualizarFranqMin As Boolean

    Call mpCarregarColecaoCobert
    mCobertSelec = ";"

    lAtualizarFranqMin = pAtualizarFranqMin
    flxCobertura.Rows = 1
    flxSelecionadas.Rows = 1

    'txtDescComercial.Text = "0"
    'txtDescExp.Text = "0"
    lCargaAux = gCarga
    gCarga = True


    'Ficha 1000757
    lCod_Clasf = Left(frmT46V101A.txtClassificacao.Tag, 4)
    lCod_Compl_Clasf = Mid(frmT46V101A.txtClassificacao.Tag, 5, 2)

    ' Plano BRFoods
    gM46V999.gEplanoSimplificado = False

    ' Verifica se possui verbas separadas
    If (frmT46V101A.cboTipCobert.ListIndex = e_TipCobert_PredioeConteudo) And (frmT46V101A.optVerbasSeparadas(0).Value = True) Then
        lPredioConteudoVerbasSeparadas = True
    Else
        lPredioConteudoVerbasSeparadas = False
    End If

    ' Calcula LMG Basico caso esteja zerado
    If val(frmT46V101A.txtISBas.Text) = 0 And gM46V999.gCod_Ramo <> e_CodRamo_Roubos Then
        lValorLMGTotal = 0
        For Each lobjstcCobert In pColecaoPedCobert
            If lobjstcCobert.gCod_Grupo_Cobert = 1 Or lobjstcCobert.gCod_Grupo_Cobert = 5 Then
                lValorLMGTotal = lValorLMGTotal + lobjstcCobert.Val_IS
            End If
        Next
        frmT46V101A.txtISBas.Text = Format(lValorLMGTotal, "standard")
    End If

    lValorLMGTotal = 0
    flxSelecionadas.Redraw = False


    'lValLMGSomaSeparado = 0

    For Each lobjstc In ColecaoCobert
        lblnAchou = False
        lValLMG = 0
        lDescComercial = 0
        lDescExperiencia = 0
        lPercent = 0
        lFranqSelec = 0
        lNum_Meses_PI = 0
        lValFatorFraqSel = 0
        lValRiscoCob = 0
        lTaxaFinal = 0
        TaxaAlterada = ""
        FranquiaAlterada = ""
        lAlterouFranquia = ""
        lAjusteEndosso = ""
        lCodGrupoCobertura = 0
        lMultiplo = 0
        lDescMultiplo = 0

        For Each lobjstcCobert In pColecaoPedCobert

            'Retirado para controle de Endosso
            'If lObjstcCobert.Cod_Cobert = lobjstc.Codigo And _
             '    lObjstcCobert.Tab_Cobert = lobjstc.Tab_Cobert Then

            'Ficha 1000757

            Call VerificaAlteracaoFranquia(lCod_Clasf, lCod_Compl_Clasf)

            If lobjstcCobert.Cod_Cobert = lobjstc.codigo Then
                lValLMG = lobjstcCobert.Val_IS
                lDescComercial = lobjstcCobert.Per_Desc_Coml
                lDescExperiencia = lobjstcCobert.Per_Desc_Exp
                lPercent = lobjstcCobert.Per_Franq
                lFranqSelec = lobjstcCobert.Val_Franq_Min
                lNum_Meses_PI = lobjstcCobert.Num_Meses_PI    'CR23794 - Syas Kit - Erro PI
                If lNum_Meses_PI > 0 And frmT46V101A.chkLMI And frmT46V101A.strCepAnt <> frmT46V101A.txtRisCEP.Text Then
                    lNum_Meses_PI = mfBuscaPerIndenitarioDefault(lobjstc.codigo)
                ElseIf frmT46V101A.chkLMI = 0 Then
                    If lNum_Meses_PI > 0 And gPiDefault = True Then
                        lNum_Meses_PI = mfBuscaPerIndenitarioDefault(lobjstc.codigo)
                    End If
                Else
                    gPiDefault = False
                End If
                lValFatorFraqSel = lobjstcCobert.Val_Fator_Franq
                lValRiscoCob = lobjstcCobert.Val_Risco_Cob
                lTaxaFinal = lobjstcCobert.Per_Taxa
                TaxaAlterada = lobjstcCobert.Alterar_Cobert

                'Ficha 100183
                FranquiaAlterada = lobjstcCobert.Alterar_Cobert
                lAlterouFranquia = lobjstcCobert.AlterouFranquia
                lAjusteEndosso = lobjstcCobert.AjusteEndosso
                'lTextoFranquia = lobjstcCobert.gTextoFranquia
                lTextoFranquia = lobjstcCobert.TextoFranquia

                lCodGrupoCobertura = lobjstcCobert.gCod_Grupo_Cobert
                If lPredioConteudoVerbasSeparadas Then
                    lComplementoDescricaoCobertura = IIf(lobjstcCobert.gCod_Grupo_Cobert = 1, DESC_COBERT_PREDIO, IIf(lobjstcCobert.Cod_Grupo_Cobertura = 5, DESC_COBERT_CONTEUDO, Empty))
                Else
                    lComplementoDescricaoCobertura = Empty
                End If

                lMultiplo = IIf(lobjstcCobert.gMultiploFraq = 0, 1, lobjstcCobert.gMultiploFraq)
                lDescMultiplo = lobjstcCobert.Per_Desc_Franq

                lblnAchou = True
                Exit For
            End If
        Next lobjstcCobert

        lObrig = False
        Select Case frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex)
            'Ficha 901282
        Case 1020, 1030, 1040
            If lDescComercial = 0 And gM46V999.gfPreenchido(frmT46V101A.txtDescComercial.Text) Then
                lDescComercial = frmT46V101A.txtDescComercial.Text
            End If
            lObrig = True
            'Ficha 901673 e 901772
        Case 2000 To 2049, 2051 To 2100, 2111 To 2138    'Ficha 1000099 - Planos 2131 a 2138 Corretor Schroeder
            If lDescComercial = 0 And gM46V999.gfPreenchido(frmT46V101A.txtDescComercial.Text) Then
                lDescComercial = frmT46V101A.txtDescComercial.Text
            End If
            If frmT46V101A.objstcPedLoc.Tip_Emissao = e_TipEmissao_RenYasuda Then
                lDescExperiencia = 0
            End If
            lObrig = True
        Case 1050
            Select Case lobjstc.codigo
            Case 38, 28
                lObrig = False
            Case Else
                lObrig = True
            End Select
        Case 1060
            Select Case lobjstc.codigo
            Case 38, 9, 22, 23, 30, 28
                lObrig = False
            Case Else
                lObrig = True
            End Select
            'Ficha 1000495 - Plano Belacity
            '''            'Ficha 901562
            '''            Case 2050, 2101, 2102
            '''                lObrig = False
        Case Else
            'If gM46V999.gEplano = True Then
             If frmT46V101A.objstcPedido.Cod_Estip >= 2000 Then
                If frmT46V101A.objstcPedido.Cod_Ramo = e_CodRamo_Empresarial _
                And (frmT46V101A.objstcPedido.Cod_Estip = 2300 _
                    Or frmT46V101A.objstcPedido.Cod_Estip = 3101 _
                    Or frmT46V101A.objstcPedido.Cod_Estip = 3095 _
                    Or frmT46V101A.objstcPedido.getPlano.EhPlanoSegmentado) Then
                    ' Não fazer nada - esses planos não tem coberturas obrigatórias
                Else
                    If lDescComercial = 0 And gM46V999.gfPreenchido(frmT46V101A.txtDescComercial.Text) Then
                        lDescComercial = frmT46V101A.txtDescComercial.Text
                    End If
                    If frmT46V101A.objstcPedLoc.Tip_Emissao = e_TipEmissao_RenYasuda Then
                        lDescExperiencia = 0
                    End If
                    lObrig = True
                End If
            End If
            ' Plano BRFoods - Inicio
            ' Obter o valor da IS fixa da cobertura (para planos com prêmio simplificado
            dValISFixa = 0
            codCobertAnt = gM46V103.lCod_Cobert
            gM46V103.lCod_Cobert = lobjstc.codigo
            If gM46V103.mfAjustesPontuais(e_Tip_Valor_ISFixa_Cobertura, TpoValor, dValISFixa, lMensagem, frmT46V101A.objstcPedLoc, 0, "") = False Then
                Call gM46V999.gpGraLog(2, lMensagem)
                'Exit Function
            End If
            gM46V103.lCod_Cobert = codCobertAnt
            If dValISFixa <> 0 Then
                lObrig = True
                gM46V999.gEplanoSimplificado = True
                AlterarCobertura = e_SimNao_Nao
                IncluirCobertura = e_SimNao_Nao
                AlterarIs = e_SimNao_Nao
            End If
            ' Plano BRFoods - Fim
            
            ' Plano Ipiranga - Inicio
            dValISLimite = 0
            codCobertAnt = gM46V103.lCod_Cobert
            gM46V103.lCod_Cobert = lobjstc.codigo
            If gM46V103.mfAjustesPontuais(e_Tip_Valor_ISLimite_Cobertura, TpoValor, dValISLimite, lMensagem, frmT46V101A.objstcPedLoc, 0, "") = False Then
                Call gM46V999.gpGraLog(2, lMensagem)
                'Exit Function
            End If
            gM46V103.lCod_Cobert = codCobertAnt
            If dValISLimite <> 0 Then
                dValISLimite = dValISLimite
            End If
            ' Plano Ipiranga - Fim
        End Select
        
        If lblnAchou = False And (lobjstc.codigo <> 1 Or (lobjstc.codigo = 1 And gM46V999.gCod_Ramo = e_CodRamo_Roubos)) And lObrig = False Then

            If lobjstc.CodGrupoCobert <> 4 Then
                flxCobertura.Rows = flxCobertura.Rows + 1
                flxCobertura.Row = flxCobertura.Rows - 1
                flxCobertura.TextMatrix(flxCobertura.Rows - 1, 0) = lobjstc.codigo
                flxCobertura.TextMatrix(flxCobertura.Rows - 1, 1) = lobjstc.Descricao
                flxCobertura.TextMatrix(flxCobertura.Rows - 1, 2) = lobjstc.Limite & "%"
                flxCobertura.TextMatrix(flxCobertura.Rows - 1, 3) = lobjstc.Tab_Cobert
                flxCobertura.TextMatrix(flxCobertura.Rows - 1, 6) = 0
                Select Case frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex)

                    '[FRH.09]->

                Case 2000 To 2049
                    flxCobertura.TextMatrix(flxCobertura.Rows - 1, 6) = 1
                    flxCobertura.RowSel = flxCobertura.Rows - 1
                    mLinhaInicial = flxCobertura.RowSel
                    mLinhaFinal = flxCobertura.RowSel
                    flxCobertura_DblClick
                    '<-[FRH.09]

                Case 1020, 1030, 1040
                    flxCobertura.TextMatrix(flxCobertura.Rows - 1, 6) = 1
                Case 1050
                    Select Case lobjstc.codigo
                    Case 38, 28
                        flxCobertura.TextMatrix(flxCobertura.Rows - 1, 6) = 0
                    Case Else
                        flxCobertura.TextMatrix(flxCobertura.Rows - 1, 6) = 1
                    End Select
                Case 1060
                    Select Case lobjstc.codigo
                    Case 38, 9, 22, 23, 30, 28
                        flxCobertura.TextMatrix(flxCobertura.Rows - 1, 6) = 0
                    Case Else
                        flxCobertura.TextMatrix(flxCobertura.Rows - 1, 6) = 1
                    End Select
                End Select
                'Ficha 1000495 - Plano Belacity
                If gM46V998.gPlanoImobiliaria Then
                    flxCobertura.TextMatrix(flxCobertura.Rows - 1, 6) = 0
                End If

                flxCobertura.TextMatrix(flxCobertura.Rows - 1, 7) = lobjstc.CodCobertLim
                flxCobertura.TextMatrix(flxCobertura.Rows - 1, 8) = lobjstc.CodGrupoCobert
            End If
        Else

            '            If lobjstc.Codigo = 1 And Not Calculando And gM46V999.gCod_Ramo <> e_CodRamo_Roubos Then
            '                If lValLMG = 0 Then
            '                    lValLMG = frmT46V101A.txtISBas.Text
            '                End If
            '                'frmT46V101A.txtISBas.Text = Format(lValLMG, "standard")
            '            End If

            If Not (lobjstc.CodGrupoCobert = 4 And (frmT46V101A.optVida(1).Value = True And gM46V111.gCod_Dizer = 0)) Then

                flxSelecionadas.Rows = flxSelecionadas.Rows + 1
                flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColCodCobert) = lobjstc.codigo
                
                'Elevador
                If lobjstc.codigo = 1103 And gM46V999.gCod_Ramo = e_CodRamo_Condominio Then
                    Select Case frmT46V101A.cboElevador.ItemData(frmT46V101A.cboElevador.ListIndex)
                    Case 0
                        flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColNomCobert) = lobjstc.Descricao
                    Case 1
                        flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColNomCobert) = lobjstc.Descricao & " COM ELEVADOR"
                    Case 2
                        flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColNomCobert) = lobjstc.Descricao & " SEM ELEVADOR"
                    End Select
                Else
                    flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColNomCobert) = lobjstc.Descricao & lComplementoDescricaoCobertura
                End If
                
                If lPredioConteudoVerbasSeparadas = False And lobjstc.CodGrupoCobert = 1 Then
                    flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColValLMG) = Format(frmT46V101A.txtISBas.Text, "standard")
                Else
                    flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColValLMG) = Format(lValLMG, "standard")
                End If

                lValorLMGTotal = lValorLMGTotal + lValLMG
                If val(flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColValLMG)) = 0 Then
                    flxSelecionadas.Col = ColValLMG
                    flxSelecionadas.Row = flxSelecionadas.Rows - 1
                    flxSelecionadas.CellBackColor = &H80000018
                Else
                    flxSelecionadas.Col = ColValLMG
                    flxSelecionadas.Row = flxSelecionadas.Rows - 1
                    flxSelecionadas.CellBackColor = vbWhite
                End If
                flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColLimiteMaximo) = 0
                flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColValRisco) = Format(lValRiscoCob, "standard")
                flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColTabCobert) = lobjstc.Tab_Cobert

                lPerMin = 0
                pTextoFranquia = Empty
                lFranqMinimo = 0
                lMinimoDias = 0
                lValFatorFraqMin = 0
                
                'Call mpGetLimFranquia(lobjstc.Codigo, lPerMin, lFranqMinimo, Dummy, dymmy1, dymmy1, Dummy1, pMuliploFraq, pDescMultFraq)

                If gM46V999.gCodProduto >= 1080 Then
                    Dim lrsFranquia As clsYasRecordSet
                    'Seta a varial de corretor pois é usada para filtrar a franquia
                    gM46V999.gCod_Corretor = val(frmT46V102A.txtCorretor(1).Tag)
                    Call gM46V999.mfFranqNovoLim(lobjstc.codigo, lValRiscoCob, frmT46V101A.objstcPedLoc.cod_clasf, frmT46V101A.objstcPedLoc.cod_compl_clasf, _
                                                    frmT46V101A.objstcPedLoc.Cod_Constr, Replace(frmT46V101A.txtRisCEP, "-", ""), frmT46V101A.cboTipCobert.ItemData(frmT46V101A.cboTipCobert.ListIndex), lNum_Meses_PI, lrsFranquia)
                    If lrsFranquia.RecordCount > 0 Then
                        lPerMin = gM46V999.gfTratarNulo(lrsFranquia("Perc_Franquia"), tpoNumerico)
                        pTextoFranquia = gM46V999.mfMontaTexofranquia(gM46V999.gfTratarNulo(lrsFranquia("Perc_Franquia"), tpoNumerico), gM46V999.gfTratarNulo(lrsFranquia("Val_Franq_Min"), tpoNumerico), gM46V999.gfTratarNulo(lrsFranquia("Dias_Franq_Min"), tpoNumerico), gM46V999.gfTratarNulo(lrsFranquia("Texto"), tpoString), gM46V999.gfTratarNulo(lrsFranquia("Compl"), tpoString), gM46V999.gfTratarNulo(lrsFranquia!Tem_VideObs, tpoString))
                        lFranqMinimo = gM46V999.gfTratarNulo(lrsFranquia("Val_Franq_Min"), tpoNumerico)
                        lMinimoDias = gM46V999.gfTratarNulo(lrsFranquia("Dias_Franq_Min"), tpoNumerico)
                        
                        '8000048888, Franquia Perda Pagamento de Aluguel
                        If lobjstcCobert.Cod_Cobert = 1117 And gM46V999.gCod_Ramo = 112 And lobjstcCobert.ObjPai.ObjPai.Cod_Produto >= 1170 Then  '1117 - Perda/Pagamento de Aluguel
                            lMinimoDias = lobjstcCobert.gValFranqMinDefault
                        End If
                        
                        lValFatorFraqMin = gM46V999.gfTratarNulo(lrsFranquia("Fator_Ajuste"), tpoNumerico)
                    End If

                Else
                    Call mpGetLimFranquia(lobjstc.codigo, lPerMin, lFranqMinimo, Dummy, dymmy1, dymmy1, Dummy1)
                End If

                lAtualizarFranqMin = pAtualizarFranqMin
                If Not lobjstcCobert Is Nothing Then
                    lAtualizarFranqMin = lAtualizarFranqMin Or ((lobjstcCobert.ObjPai.ObjPai.Num_Apol_An > 0 Or lobjstcCobert.ObjPai.ObjPai.objstcCtrlEmis.Num_Apol_Mar > 0) And lAlterouFranquia <> "SIM" And lobjstcCobert.gMultiploFraq = 0 And lobjstcCobert.ObjPai.ObjPai.Cod_Produto >= 1080 And gM46V111.gCod_Dizer = 0)
                End If
                
                lAtualizarFranqMin = pCargaNormal And lAtualizarFranqMin
                
                'If pAtualizarFranqMin Or lNum_Meses_PI > 0 Then
                If lAtualizarFranqMin Then
                    flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColPerFranq) = lPerMin
                    If lMinimoDias = 0 Then
                        flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColTipFranq) = "R$"
                        flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColValFranq) = Format(lFranqMinimo, "standard")
                        flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColValFranMin) = Format(lFranqMinimo, "standard")
                    Else
                        flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColTipFranq) = "Dias"
                        flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColValFranq) = lMinimoDias
                        flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColValFranMin) = lMinimoDias
                    End If
                    flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColMultiploFranquia) = "1x"
                    flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColDesconto) = Format(pDescMultFraq, "standard")
                    flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColTextoFranquia) = pTextoFranquia
                Else
                    flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColPerFranq) = lPercent
                    If lMinimoDias = 0 Then
                        flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColTipFranq) = "R$"
                    flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColValFranq) = Format(lFranqSelec, "standard")
                        flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColValFranMin) = Format(lFranqMinimo, "standard")
                    Else
                        flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColTipFranq) = "Dias"
                        flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColValFranq) = lFranqSelec
                        flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColValFranMin) = lMinimoDias
                        End If
                    
                    If lFranqMinimo > 0 Or lMinimoDias > 0 Then
                        If (lFranqSelec / IIf(lFranqMinimo = 0, lMinimoDias, lFranqMinimo)) > 1 And gM46V111.gCod_Dizer > 0 Then
                            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColMultiploFranquia) = "3x"
                        Else
                            flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColMultiploFranquia) = lMultiplo & "x"
                    End If
                    Else
                    flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColMultiploFranquia) = lMultiplo & "x"
                    End If
                    flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColDesconto) = Format(lDescMultiplo, "standard")
                    flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColTextoFranquia) = lTextoFranquia
                End If

                If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda And gM46V999.gCod_Ramo = e_CodRamo_Condominio Then
                    'flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColValFranq) = Format(lFranqMinimo, "standard")
                End If

                flxSelecionadas.Redraw = True
                flxSelecionadas.Col = ColValFranMin
                flxSelecionadas.CellBackColor = vbYellow
                If gM46V999.gCodProduto < 1080 Then
                    flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColValFranMin) = Format(lFranqMinimo, "standard")
                End If
                flxSelecionadas.Col = ColValFranMin
                flxSelecionadas.CellBackColor = vbYellow

                If val(flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColValFranq)) = 0 Then
                    flxSelecionadas.Col = ColValFranq
                    If gM46O001.PermiteAlterarFranquia Then flxSelecionadas.CellBackColor = &H80000018
                Else
                    flxSelecionadas.Col = ColValFranq
                    flxSelecionadas.CellBackColor = vbWhite
                End If

                'Ficha 1000183
                If gM46V999.gCod_Ramo = e_CodRamo_Roubos Or gM46V999.gNovaPropostaLMI Or gM46O001.PermiteAlterarFranquia Then
                    flxSelecionadas.Col = ColPerFranq
                    flxSelecionadas.CellBackColor = &H80000018
                End If
                
                If gM46V999.gCod_Ramo <> e_CodRamo_Roubos Then
                    flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColPInd) = lNum_Meses_PI
                    flxSelecionadas.Col = ColPInd
                    If gM46O001.PermiteAlterarFranquia Then flxSelecionadas.CellBackColor = &H80000018
                End If

                flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColDescComl) = Format(lDescComercial, "standard")

                'Desconto Comercial na Tela Principal
                If Not Calculando Then
                    If gM46V999.gTipoAcesso = DIGITACAO And lDescComercial = 0 And frmT46V101A.objstcPedido.DAT_DIGIT = 0 And frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex) < 2000 Then
                        '*Alterado por Luiz Alberto em 27/01/2016 - para que não seja zerado o desconto comercial (cobertura vida), ao carregar a aplicação
                        '*Não alterei a linha quando: lDescComercial = 0, apenas evito a atribuição como o default é: 0,00
                        '*frmT46V101A.txtDescComercial = "0,00"
                        '****************************************
                        If Format(lDescComercial, "standard") > 0 Then frmT46V101A.txtDescComercial = Format(lDescComercial, "standard")
                    Else
                        If gCarga Then    '[FRH.09]
                            If gM46V111.gCod_plano < 2000 Then
                                '*Alterado por Luiz Alberto em 27/01/2016
                                '*frmT46V101A.txtDescComercial = Format(lDescComercial, "standard")
                                '****************************************
                                If Format(lDescComercial, "standard") > 0 Then frmT46V101A.txtDescComercial = Format(lDescComercial, "standard")
                            End If
                        End If
                    End If
                End If
                If val(flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColDescComl)) = 0 Then
                    flxSelecionadas.Col = ColDescComl
                    flxSelecionadas.Row = flxSelecionadas.Rows - 1
                    flxSelecionadas.CellBackColor = &H80000018
                Else
                    flxSelecionadas.Col = ColDescComl
                    flxSelecionadas.Row = flxSelecionadas.Rows - 1
                    flxSelecionadas.CellBackColor = vbWhite
                End If
                flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColDescExp) = Format(lDescExperiencia, "standard")
                If val(flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColDescExp)) = 0 Then
                    flxSelecionadas.Col = ColDescExp
                    flxSelecionadas.Row = flxSelecionadas.Rows - 1
                    flxSelecionadas.CellBackColor = &H80000018
                Else
                    flxSelecionadas.Col = ColDescExp
                    flxSelecionadas.Row = flxSelecionadas.Rows - 1
                    flxSelecionadas.CellBackColor = vbWhite
                End If
                flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColCobObrigat) = 0
                Select Case frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex)
                Case 1020, 1030, 1040, Is > 2000
                    If gM46V998.gPlanoImobiliaria = True _
                    Or (frmT46V101A.objstcPedido.Cod_Ramo = e_CodRamo_Empresarial And frmT46V101A.objstcPedido.Cod_Estip = 2300) _
                    Or (frmT46V101A.objstcPedido.Cod_Ramo = e_CodRamo_Empresarial And frmT46V101A.objstcPedido.Cod_Estip = 3101) _
                    Or frmT46V101A.objstcPedido.Cod_Estip = 3095 _
                    Or (frmT46V101A.objstcPedido.Cod_Estip >= 2228 And frmT46V101A.objstcPedido.Cod_Estip <= 2271) _
                    Or (frmT46V101A.objstcPedido.getPlano.EhPlanoSegmentado) _
                    Then
                        flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColCobObrigat) = 0
                    Else
                        flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColCobObrigat) = 1
                    End If
                Case 1050
                    Select Case lobjstc.codigo
                    Case 38, 28
                        flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColCobObrigat) = 0
                    Case Else
                        flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColCobObrigat) = 1
                    End Select
                Case 1060
                    Select Case lobjstc.codigo
                    Case 38, 9, 22, 23, 30, 28
                        flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColCobObrigat) = 0
                    Case Else
                        flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColCobObrigat) = 1
                    End Select

                End Select
                flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColPerLimite) = lobjstc.Limite
                If mfBuscaPerIndenitario(lobjstc.codigo) Then
                    'Os PIs foram passados para a tabela de PI padrão TBM_Per_Ind_Padrao
                    '1200539 - Tarifa 1038 e 1039 Junho 2012 - Item 11
                    '                If (frmT46V101A.objstcPedido.Cod_Produto >= 1038) And _
                                     '                   (frmT46V101A.objstcPedido.Cod_Ramo = e_CodRamo_Empresarial) And _
                                     '                   (frmT46V101A.objstcPedido.Cod_Estip >= 2170 And frmT46V101A.objstcPedido.Cod_Estip <= 2177) And _
                                     '                   (lobjstc.Codigo = 39) Then
                    '                        lNum_Meses_PI = 2
                    '                End If
                    '
                    '                '1201120 - Tarifa 1052 e 1053 - Item 3
                    '                If (frmT46V101A.objstcPedido.Cod_Produto >= 1052) And _
                                     '                   (frmT46V101A.objstcPedido.Cod_Ramo = e_CodRamo_Empresarial) And _
                                     '                   (frmT46V101A.objstcPedido.Cod_Estip >= 2194 And frmT46V101A.objstcPedido.Cod_Estip <= 2201) And _
                                     '                   (lobjstc.Codigo = 39) Then
                    '                        lNum_Meses_PI = 2
                    '                End If
                    '
                    '                '1200806 - Tarifa 1044 e 1045 Setembro 2012 - Item 26
                    '                If (frmT46V101A.objstcPedido.Cod_Produto >= 1044) And _
                                     '                   (frmT46V101A.objstcPedido.Cod_Ramo = e_CodRamo_Empresarial) And _
                                     '                   (frmT46V101A.objstcPedido.Cod_Estip >= 2183 And frmT46V101A.objstcPedido.Cod_Estip <= 2184) And _
                                     '                   (lobjstc.Codigo = 3 Or lobjstc.Codigo = 39) Then
                    '                        lNum_Meses_PI = 3
                    '                End If
                    '
                    '                '1200876 - Planos REPOM
                    '                If (frmT46V101A.objstcPedido.Cod_Produto >= 1044) And _
                                     '                   (frmT46V101A.objstcPedido.Cod_Ramo = e_CodRamo_Empresarial) And _
                                     '                   (frmT46V101A.objstcPedido.Cod_Estip >= 2185 And frmT46V101A.objstcPedido.Cod_Estip <= 2193) And _
                                     '                   (lobjstc.Codigo = 3) Then
                    '                        lNum_Meses_PI = 3
                    '                End If

                    flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColPermitePInd) = True
                    If lNum_Meses_PI = 0 Then
                        lNum_Meses_PI = mfBuscaPerIndenitarioDefault(lobjstc.codigo)
                    End If
                    flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColPInd) = IIf(lNum_Meses_PI = 0, 12, lNum_Meses_PI)
                    flxSelecionadas.Col = ColPInd
                    flxSelecionadas.CellBackColor = vbWhite
                End If
                lNum_Meses_PI = 0   'zera variável porque estava carregando errado em calculos já salvos (ao trocar planos)

                'Supermercado SMART
                If (frmT46V101A.objstcPedido.Cod_Produto >= 1042) And _
                   (frmT46V101A.objstcPedido.Cod_Ramo = e_CodRamo_Empresarial) And _
                   (frmT46V101A.objstcPedido.Cod_Estip >= 2178 And frmT46V101A.objstcPedido.Cod_Estip <= 2182) And _
                   (lobjstc.codigo = 39 Or lobjstc.codigo = 3) Then
                    lNum_Meses_PI = 3
                    flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColPermitePInd) = True
                    flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColPInd) = IIf(lNum_Meses_PI = 0, 12, lNum_Meses_PI)
                    flxSelecionadas.Col = ColPInd
                    flxSelecionadas.CellBackColor = vbWhite
                End If

                If mfBuscaLucrosCessantes(lobjstc.codigo) Then
                    flxSelecionadas.Col = ColValRisco
                    flxSelecionadas.CellBackColor = vbWhite
                End If
                flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColCoefFranq) = lValFatorFraqSel
                flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColTaxaFinal) = Format(lTaxaFinal, "0.00000")
                flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColTaxaAlterada) = TaxaAlterada
                flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColCobertPerLimite) = lobjstc.CodCobertLim
                If TaxaAlterada = "ALTERADO" Then
                    flxSelecionadas.Col = ColTaxaFinal
                    flxSelecionadas.CellForeColor = vbRed
                End If

                flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColAlterouFranquia) = lAlterouFranquia
                If lAlterouFranquia = "SIM" Then
                    flxSelecionadas.Col = ColAlterouFranquia
                    flxSelecionadas.CellForeColor = vbRed
                End If
                
                flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColAjusteEndosso) = lAjusteEndosso
                flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColValISAte) = mfRetornaLMG(flxSelecionadas.Rows - 1)
                flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColValISmin) = mfRetornaLMGMin(flxSelecionadas.Rows - 1)
                flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColCodGrupoCobert) = lobjstc.CodGrupoCobert
                Call mpCalcularLimite(flxSelecionadas.Rows - 1)
                mCobertSelec = mCobertSelec & flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColCodCobert) & ";"

                ' Plano BRFoods - Inicio
                If dValISFixa <> 0 Then
                    'flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColLimiteMaximo) = Format(dValISFixa, "standard")
                    flxSelecionadas.Col = ColValLMG
                    flxSelecionadas.Row = flxSelecionadas.Rows - 1
                    flxSelecionadas.CellBackColor = &H80000018

                    flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColValLMG) = Format(dValISFixa, "standard")

                    flxSelecionadas.Col = ColValRisco
                    flxSelecionadas.Row = flxSelecionadas.Rows - 1
                    flxSelecionadas.CellBackColor = &H80000018

                    flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColValRisco) = Format(dValISFixa, "standard")
                End If
                ' Plano BRFoods - Fim

                If dValISLimite <> 0 Then
                    flxSelecionadas.Col = ColLimiteMaximo
                    flxSelecionadas.TextMatrix(flxSelecionadas.Rows - 1, ColLimiteMaximo) = Format(dValISLimite, "standard")
                End If

            End If    'Verifica se vida está habilitado

        End If
    Next lobjstc


    '    If Not Calculando And gM46V999.gCod_Ramo <> e_CodRamo_Roubos Then
    '        If lValLMGSomaSeparado = 0 Then
    '            lValLMGSomaSeparado = frmT46V101A.txtISBas.Text
    '        End If
    '        frmT46V101A.txtISBas.Text = Format(lValLMGSomaSeparado, "standard")
    '    End If
    '
    If gM46V999.gCod_Ramo = e_CodRamo_Roubos Then
        frmT46V101A.txtISBas.Text = Format(lValorLMGTotal, "standard")
    End If

    'LMG(LC)
    Call mpRecalcularLMGLC

    gM46V999.gChamaSelecionarCoberturas = False
    gAltDadosCobert = False
    flxSelecionadas.Row = 0
    flxSelecionadas.ColSel = ColValLMG
    flxSelecionadas_SelChange

    If flxSelecionadas.Row > 0 Then
        flxSelecionadas.Row = 1
    End If

    flxSelecionadas.Redraw = True
    flxSelecionadas.Refresh
    gCarga = lCargaAux
    'CR23794 - Syas Kit - Erro PI
    gPiDefault = False
    
    Call mpVerificarLimiteDependente
End Sub

Private Sub mpCarregarColecaoCobert()
    Dim lobjstc As stcA46V101B                     'Objeto da classe de estrutura stcA46V101A.
    Dim lbdTAB_Cobert As ADODB.Connection                     'Arquivo: P00Tarifa.
    Dim lMensagem As String                     'Mensagem.
    Dim lrsTab_Cobert As clsYasRecordSet                     'Registro : TAB_COBERT.
    Dim lPredioConteudoVerbasSeparadas As Boolean     'Prédio e conteúdo com verbas separadas.
    Dim lSelect As String
    Dim lCoberturaRestrita As Integer

    Set ColecaoCobert = New Collection

    If (frmT46V101A.cboTipCobert.ListIndex = e_TipCobert_PredioeConteudo) And (frmT46V101A.optVerbasSeparadas(0).Value = True) Then
        lPredioConteudoVerbasSeparadas = True
    Else
        lPredioConteudoVerbasSeparadas = False
    End If

    'Nova tarifa 1092 / 1093
    If val(frmT46V102A.txtCorretor(1).Tag) = 908624 And frmT46V101A.objstcPedido.Cod_Produto >= 1092 Then
        lCoberturaRestrita = True
    Else
        lCoberturaRestrita = False
    End If

    If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTAB_Cobert, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If

    lSelect = " Select distinct DSC_Cobert, Cod_Cobert, Per_Lim_Max_Cobert,TAB_COBERT  , PER_FRANQ , VAL_FRANQ_MIN , COD_LIM_COBERT, NUM_ORDEM_COBERT, COD_GRUPO_COBERT from TAB_COBERT"
    lSelect = lSelect & " Where cod_Ramo = " & gM46V999.gCod_Ramo
    If val(Right(frmT46V101A.txtRisRegiao.Tag, 4)) <> 0 And gM46V999.gCodProduto <= gM46V999.TarifaAntiga Then
        lSelect = lSelect & " and tab_cobert = " & val(Right(frmT46V101A.txtRisRegiao.Tag, 4))
    Else
        lSelect = lSelect & " and tab_cobert = " & frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex)
    End If
    lSelect = lSelect & " and COD_TARIF_INI <= " & gM46V999.gCodProduto
    lSelect = lSelect & " and COD_TARIF_TER >= " & gM46V999.gCodProduto

    If lPredioConteudoVerbasSeparadas = False Then
        lSelect = lSelect & " and COD_GRUPO_COBERT in (1, 2, 3, 4)"
    Else
        lSelect = lSelect & " and COD_GRUPO_COBERT in (1, 2, 3, 4, 5)"
    End If

    'Nova tarifa 1092 / 1093
    If lCoberturaRestrita Then
        lSelect = lSelect & " and (cod_restrito IN(1,0,2)  or cod_restrito is null)"
    Else
        'Residencial Supremo
        If frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex) = 30 Then
            lSelect = lSelect & " and (cod_restrito IN(0,2)  or cod_restrito is null)"
        Else
        lSelect = lSelect & " and cod_cobert <> 1215"
    End If
    End If

    If gM46V999.gCodProduto >= 1122 Then '11.22
        gM46V999.gNum_Apolice = frmT46V101A.objstcPedido.Num_Apol_An
        If Not (frmT46V101A.objstcPedLoc.Cod_DizerItem >= e_CodDizerInclusao And gM46V999.gTarifa_Apolice > 0 And gM46V999.gTarifa_Apolice < 1122) Then
            lSelect = lSelect & " and NOT (Cod_ramo = 114 AND cod_cobert = 9636 AND Tab_Cobert = 11)"
        End If
    End If
    
    lSelect = lSelect & " order by NUM_ORDEM_COBERT, Cod_Cobert"

    If gM46V999.gfObtRegistro(lbdTAB_Cobert, lSelect, lrsTab_Cobert, lMensagem) = False Then
        Call gM46V999.gpFecha2(lbdTAB_Cobert)
        Call gM46V999.gpGraLog(2, lMensagem)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Set ColecaoCobert = New Collection

    Do While Not lrsTab_Cobert.EOF
        Set lobjstc = New stcA46V101B
        lobjstc.Descricao = Trim(lrsTab_Cobert("DSC_Cobert"))
        lobjstc.codigo = lrsTab_Cobert("Cod_Cobert")
        lobjstc.Limite = lrsTab_Cobert("Per_Lim_Max_Cobert")
        lobjstc.Tab_Cobert = lrsTab_Cobert("TAB_COBERT")
        lobjstc.Per_Franq = lrsTab_Cobert("PER_FRANQ")
        lobjstc.Val_Franq_Min = lrsTab_Cobert("VAL_FRANQ_MIN")
        lobjstc.CodCobertLim = lrsTab_Cobert("COD_LIM_COBERT")
        lobjstc.CodGrupoCobert = lrsTab_Cobert("COD_GRUPO_COBERT")
        ColecaoCobert.Add lobjstc
        lrsTab_Cobert.MoveNext
    Loop
End Sub

Public Sub mpCarregarGridCoberturas()
    Dim lobjstc As stcA46V101B                     'Objeto da classe de estrutura stcA46V101A.
    Dim lLinha As Long                     'Linhas.
    Dim lCargaTMP As Boolean

    '    Dim lSepararVerbas                      As Boolean          'Verbas foram separadas
    '    Separar Verbas
    '    lSepararVerbas = True
    '
    '    For Each lobjstcCobert In pColecaoPedCobert         'Verifica se existe a cobertura de verbas separadas na cobertura
    '        If lobjstcCobert.Cod_Grupo_Cobertura = 5 Then
    '            lSepararVerbas = False
    '            Exit For
    '        End If
    '    Next lobjstcCobert

    If frmT46V101A.cboPlano.ListIndex <> -1 Then
        lCargaTMP = gCarga
        gCarga = True
        lLinha = 1

        'Carrega Coleção de Coberturas
        Call mpCarregarColecaoCobert

        'Ficha 901562
        If frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex) > 1010 _
        And gM46V998.gPlanoImobiliaria = False _
        And Not (frmT46V101A.objstcPedido.getPlano.EhPlanoSegmentado) _
        Then    'Ficha 1000495 - Plano Belacity
            For Each lobjstc In ColecaoCobert
                flxSelecionadas.Rows = lLinha + 1
                flxSelecionadas.TextMatrix(lLinha, ColCodCobert) = lobjstc.codigo
                flxSelecionadas.TextMatrix(lLinha, ColNomCobert) = lobjstc.Descricao
                'Código da Básica
                If lobjstc.codigo = 1 Then
                    flxSelecionadas.TextMatrix(lLinha, ColValLMG) = Format(frmT46V101A.txtISBas.Text, "standard")
                Else
                    flxSelecionadas.TextMatrix(lLinha, ColValLMG) = "0,00"
                End If
                flxSelecionadas.Col = ColValLMG
                flxSelecionadas.Row = flxSelecionadas.Rows - 1
                flxSelecionadas.CellBackColor = &H80000018
                flxSelecionadas.TextMatrix(lLinha, ColLimiteMaximo) = 0
                flxSelecionadas.TextMatrix(lLinha, ColTabCobert) = lobjstc.Tab_Cobert
                flxSelecionadas.TextMatrix(lLinha, ColPerFranq) = lobjstc.Per_Franq
                flxSelecionadas.TextMatrix(lLinha, ColValFranq) = lobjstc.Val_Franq_Min
                flxSelecionadas.Col = ColValFranq
                flxSelecionadas.Row = flxSelecionadas.Rows - 1
                flxSelecionadas.CellBackColor = &H80000018
                flxSelecionadas.TextMatrix(lLinha, ColDescComl) = "0,00"
                flxSelecionadas.Col = ColDescComl
                flxSelecionadas.Row = flxSelecionadas.Rows - 1
                flxSelecionadas.CellBackColor = &H80000018
                flxSelecionadas.TextMatrix(lLinha, ColDescExp) = "0,00"
                flxSelecionadas.Col = ColDescExp
                flxSelecionadas.Row = flxSelecionadas.Rows - 1
                flxSelecionadas.CellBackColor = &H80000018
                flxSelecionadas.TextMatrix(lLinha, ColCobObrigat) = 1
                Select Case frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex)
                Case 1020, 1030, 1040
                Case 1050
                    Select Case lobjstc.codigo
                    Case 38, 28
                        flxSelecionadas.TextMatrix(lLinha, ColCobObrigat) = 0
                    End Select
                Case 1060
                    Select Case lobjstc.codigo
                    Case 22, 23, 30, 28, 9, 38
                        flxSelecionadas.TextMatrix(lLinha, ColCobObrigat) = 0
                    End Select
                End Select
            
                If mfBuscaPerIndenitario(lobjstc.codigo) Then
                    Dim intPeriodoIndenitarioDefault As Integer
                    intPeriodoIndenitarioDefault = mfBuscaPerIndenitarioDefault(lobjstc.codigo)
                    'Buscar Período Indenitário Default

                    flxSelecionadas.TextMatrix(lLinha, ColPermitePInd) = True
                    flxSelecionadas.TextMatrix(lLinha, ColPInd) = intPeriodoIndenitarioDefault
                    flxSelecionadas.Col = ColPInd
                    flxSelecionadas.CellBackColor = vbWhite
                Else
                    flxSelecionadas.TextMatrix(lLinha, ColPermitePInd) = False
                    flxSelecionadas.TextMatrix(lLinha, ColPInd) = 0
                End If

                flxSelecionadas.TextMatrix(lLinha, ColPerLimite) = lobjstc.Limite
                flxSelecionadas.TextMatrix(lLinha, ColValISAte) = mfRetornaLMG(lLinha)
                flxSelecionadas.TextMatrix(lLinha, ColValISmin) = mfRetornaLMGMin(lLinha)
                flxSelecionadas.TextMatrix(lLinha, ColCodGrupoCobert) = lobjstc.CodGrupoCobert
                Call mpCalcularLimite(CInt(lLinha))
                flxSelecionadas.TextMatrix(lLinha, ColValRisco) = flxSelecionadas.TextMatrix(lLinha, ColLimiteMaximo)
                flxSelecionadas.TextMatrix(lLinha, ColCobertPerLimite) = lobjstc.CodCobertLim
                flxSelecionadas.TextMatrix(lLinha, ColCodGrupoCobert) = lobjstc.CodGrupoCobert
                flxSelecionadas.TextMatrix(lLinha, ColMultiploFranquia) = "1x"
                lLinha = lLinha + 1
            Next lobjstc
        Else
            For Each lobjstc In ColecaoCobert
                If lobjstc.codigo = 1 And gM46V999.gCod_Ramo <> e_CodRamo_Roubos Then
                    flxSelecionadas.Rows = 2
                    flxSelecionadas.TextMatrix(1, ColCodCobert) = lobjstc.codigo
                    flxSelecionadas.TextMatrix(1, ColNomCobert) = lobjstc.Descricao
                    flxSelecionadas.TextMatrix(1, ColValLMG) = Format(frmT46V101A.txtISBas.Text, "standard")
                    flxSelecionadas.TextMatrix(1, ColValRisco) = Format(frmT46V101A.txtValoRisco.Text, "standard")
                    flxSelecionadas.TextMatrix(1, ColLimiteMaximo) = lobjstc.Limite & "%"
                    flxSelecionadas.TextMatrix(1, ColTabCobert) = lobjstc.Tab_Cobert
                    flxSelecionadas.TextMatrix(1, ColPerFranq) = "0,00"
                    flxSelecionadas.TextMatrix(1, ColValFranq) = "0,00"
                    If gM46V999.gCod_Ramo = e_CodRamo_Condominio Then
                        'flxSelecionadas.TextMatrix(lLinha, ColPerFranq) = lobjstc.Per_Franq
                        'flxSelecionadas.TextMatrix(lLinha, ColValFranq) = lobjstc.Val_Franq_Min
                    End If
                    flxSelecionadas.TextMatrix(1, ColCobObrigat) = 0
                    flxSelecionadas.TextMatrix(1, ColCobertPerLimite) = lobjstc.CodCobertLim
                    flxSelecionadas.TextMatrix(1, ColCodGrupoCobert) = lobjstc.CodGrupoCobert
                    flxSelecionadas.TextMatrix(1, ColMultiploFranquia) = "1x"
                Else
                    If lobjstc.CodGrupoCobert <> 4 Then
                        flxCobertura.Rows = lLinha + 1
                        flxCobertura.TextMatrix(lLinha, 0) = lobjstc.codigo
                        flxCobertura.TextMatrix(lLinha, 1) = lobjstc.Descricao
                        flxCobertura.TextMatrix(lLinha, 2) = lobjstc.Limite & "%"
                        flxCobertura.TextMatrix(lLinha, 3) = lobjstc.Tab_Cobert
                        flxCobertura.TextMatrix(lLinha, 4) = lobjstc.Per_Franq
                        flxCobertura.TextMatrix(lLinha, 5) = lobjstc.Val_Franq_Min
                        flxCobertura.TextMatrix(lLinha, 6) = 0
                        flxCobertura.TextMatrix(lLinha, 7) = lobjstc.CodCobertLim
                        flxCobertura.TextMatrix(lLinha, 8) = lobjstc.CodGrupoCobert
                        lLinha = lLinha + 1
                    End If
                End If
            Next lobjstc
        End If

        gCarga = lCargaTMP
    End If
End Sub

Public Sub mpCalcularLimite(pLinhas As Long)
    Dim lI As Long
    Dim lValorIS As Double
    
    'Não executa se for planinho da 1080 pra cima
    'If gM46V999.gCodProduto >= 1080 And gM46V111.gCod_plano >= 2000 Then Exit Sub

    '''******* OBSERVAÇÃO : Nos cálculo de limite , está sendo dividido duas vezes por 100
    '''                     e depois multiplicado por 100 novamente para poder arredondar
    '''                     os valores (motivo principal : Planos da correta).
    If frmT46V101A.txtISBas.Text = "" Then Exit Sub
    lValorIS = CDbl(frmT46V101A.txtISBas.Text)

    With flxSelecionadas
        If pLinhas = 0 Then
            For lI = 1 To .Rows - 1
                If gM46V999.CDblx(.TextMatrix(lI, ColPerLimite)) = 0 Then
                    .TextMatrix(lI, ColLimiteMaximo) = Format(lValorIS, "standard")
                Else
                    If Not gM46V999.gfPreenchido(.TextMatrix(lI, ColPerLimite)) Then
                        .TextMatrix(lI, ColPerLimite) = 0
                    End If
                    'Ficha 1000427
                    If gM46V999.gCodProduto > 987 Then
                        If gM46V111.gCod_plano > 2000 Then    'planinho tem que arredondar
                            .TextMatrix(lI, ColLimiteMaximo) = Format(Round((lValorIS * .TextMatrix(lI, ColPerLimite) / 100) / 10) * 10, "standard")
                        Else
                            .TextMatrix(lI, ColLimiteMaximo) = Format(((lValorIS * .TextMatrix(lI, ColPerLimite) / 100) / 10) * 10, "standard")
                        End If
                    Else
                        .TextMatrix(lI, ColLimiteMaximo) = Format(Round((lValorIS * .TextMatrix(lI, ColPerLimite) / 100) / 10) * 10, "standard")
                    End If
                End If
                If gM46V999.CDblx(.TextMatrix(lI, ColValISAte)) = 0 Then .TextMatrix(lI, ColValISAte) = "9999999999999"
                If CDbl(.TextMatrix(lI, ColValISAte)) < CDbl(.TextMatrix(lI, ColLimiteMaximo)) Then
                    .TextMatrix(lI, ColLimiteMaximo) = Format(.TextMatrix(lI, ColValISAte), "standard")
                End If
                
                If gM46V999.gCodProduto >= 1080 Then
                If gM46V999.CDblx(.TextMatrix(lI, ColValISmin)) = 0 Then .TextMatrix(lI, ColValISmin) = 0
                If CDbl(.TextMatrix(lI, ColValISmin)) > CDbl(.TextMatrix(lI, ColLimiteMaximo)) Then
                    .TextMatrix(lI, ColLimiteMaximo) = Format(.TextMatrix(lI, ColValISmin), "standard")
                End If
                .TextMatrix(lI, ColLimiteMaximo) = Format(AjustePontualValorMaximo3002(.TextMatrix(lI, ColCodCobert), .TextMatrix(lI, ColLimiteMaximo), .TextMatrix(lI, ColValISmin)), "standard")
                End If
                
                If .TextMatrix(lI, ColCobObrigat) = "1" Then
                    .TextMatrix(lI, ColValLMG) = .TextMatrix(lI, ColLimiteMaximo)
                End If
            Next lI
        Else
            If .TextMatrix(pLinhas, ColPerLimite) = 0 Then
                .TextMatrix(pLinhas, ColLimiteMaximo) = Format(lValorIS, "standard")
            Else
                If Not gM46V999.gfPreenchido(.TextMatrix(pLinhas, ColPerLimite)) Then
                    .TextMatrix(lI, ColPerLimite) = 0
                End If
                'Ficha 1000427
                If gM46V999.gCodProduto > 987 Then
                    If gM46V111.gCod_plano > 2000 Then    'planinho tem que arredondar
                        .TextMatrix(pLinhas, ColLimiteMaximo) = Format(Round((lValorIS * .TextMatrix(pLinhas, ColPerLimite) / 100) / 10) * 10, "standard")
                    Else
                        .TextMatrix(pLinhas, ColLimiteMaximo) = Format(((lValorIS * .TextMatrix(pLinhas, ColPerLimite) / 100) / 10) * 10, "standard")
                    End If
                Else
                    .TextMatrix(pLinhas, ColLimiteMaximo) = Format(Round((lValorIS * .TextMatrix(pLinhas, ColPerLimite) / 100) / 10) * 10, "standard")
                End If
            End If
            If .TextMatrix(pLinhas, ColValISAte) = "" Then .TextMatrix(pLinhas, ColValISAte) = "9999999999999999"
            If CDbl(.TextMatrix(pLinhas, ColValISAte)) > 0 And CDbl(.TextMatrix(pLinhas, ColValISAte)) < CDbl(.TextMatrix(pLinhas, ColLimiteMaximo)) And gM46V111.gCod_plano < 2000 Then
                .TextMatrix(pLinhas, ColLimiteMaximo) = Format(.TextMatrix(pLinhas, ColValISAte), "standard")
            End If
            
            If gM46V999.gCodProduto >= 1080 Then
                If .TextMatrix(pLinhas, ColValISmin) = "" Then .TextMatrix(pLinhas, ColValISmin) = "0"
                If CDbl(.TextMatrix(pLinhas, ColValISmin)) > CDbl(.TextMatrix(pLinhas, ColLimiteMaximo)) Then
                    .TextMatrix(pLinhas, ColLimiteMaximo) = Format(.TextMatrix(pLinhas, ColValISmin), "standard")
                End If
                .TextMatrix(pLinhas, ColLimiteMaximo) = Format(AjustePontualValorMaximo3002(.TextMatrix(pLinhas, ColCodCobert), .TextMatrix(pLinhas, ColLimiteMaximo), .TextMatrix(pLinhas, ColValISmin)), "standard")
            End If
            
            If .TextMatrix(pLinhas, ColCobObrigat) = "1" Then
                .TextMatrix(pLinhas, ColValLMG) = .TextMatrix(pLinhas, ColLimiteMaximo)
            End If
        End If
    End With


    If gCarga = False Then
        Call mpVerificarLimiteDependente
    End If
End Sub
Private Sub txtDescExp_Change()
    Call gpLimPremios
End Sub

Private Sub txtDescExp_GotFocus()
    With txtDescExp
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtDescExp_LostFocus()
    With txtDescExp
        If gM46V999.gfPreenchido(.Text) Then
            If IsNumeric(.Text) Then
                .Text = Format(.Text, "standard")
            Else
                .Text = Format(0, "standard")
            End If
        Else
            .Text = Format(0, "standard")
        End If
        If val(.Text) >= 100 Then
            .Text = "99,99"
        End If
    End With
End Sub

Private Sub mpVerificarLimiteDependente()
    Dim lI As Long
    Dim lFind As Long
    Dim lValorIS As Double
    Dim dValISLimite As Double
    Dim codCobertAnt As Integer
    Dim lMensagem As String
    
    'Não executa se for planinho da 1080 pra cima
    'If gM46V999.gCodProduto >= 1080 And gM46V111.gCod_plano >= 2000 Then Exit Sub

    If frmT46V101A.txtISBas.Text = "" Then Exit Sub

    lValorIS = CDbl(frmT46V101A.txtISBas.Text)
    With flxSelecionadas
        For lI = 1 To .Rows - 1
            ' Plano Ipiranga - Inicio
            dValISLimite = 0
            codCobertAnt = gM46V103.lCod_Cobert
            gM46V103.lCod_Cobert = .TextMatrix(lI, ColCodCobert)
            If gM46V103.mfAjustesPontuais(e_Tip_Valor_ISLimite_Cobertura, TpoValor, dValISLimite, lMensagem, frmT46V101A.objstcPedLoc, 0, "") = False Then
                Call gM46V999.gpGraLog(2, lMensagem)
                'Exit Function
            End If
            gM46V103.lCod_Cobert = codCobertAnt
            If dValISLimite <> 0 Then
                dValISLimite = dValISLimite
            End If
            ' Plano Ipiranga - Fim
        
            If (.TextMatrix(lI, ColCodCobert) = "1") Or (InStr(UCase(.TextMatrix(lI, ColNomCobert)), DESC_COBERT_CONTEUDO)) Then
                .TextMatrix(lI, ColLimiteMaximo) = Format(lValorIS, "standard")
            Else
                If .TextMatrix(lI, ColCobertPerLimite) = "1" Then
                    ' Plano Ipiranga - Inicio
                    If dValISLimite <> 0 Then
                        .TextMatrix(lI, ColLimiteMaximo) = Format(dValISLimite, "standard")
                    Else
                    ' Plano Ipiranga - Fim
                    'Ficha 1000427
                    If gM46V999.gCodProduto > 987 Then
                        If gM46V111.gCod_plano > 2000 Then    'planinho tem que arredondar
                            .TextMatrix(lI, ColLimiteMaximo) = Format(Round((lValorIS * .TextMatrix(lI, ColPerLimite) / 100) / 10) * 10, "standard")
                        Else
                            .TextMatrix(lI, ColLimiteMaximo) = Format(((lValorIS * .TextMatrix(lI, ColPerLimite) / 100) / 10) * 10, "standard")
                        End If
                    Else
                        .TextMatrix(lI, ColLimiteMaximo) = Format(Round((lValorIS * .TextMatrix(lI, ColPerLimite) / 100) / 10) * 10, "standard")
                    End If
                    ' Plano Ipiranga - Inicio
                    End If
                    ' Plano Ipiranga - Fim
                Else
                    If InStr(1, mCobertSelec, ";" & .TextMatrix(lI, ColCobertPerLimite) & ";") > 0 Then
                        For lFind = 1 To .Rows - 1
                            If .TextMatrix(lFind, ColCodCobert) = .TextMatrix(lI, ColCobertPerLimite) Then
                                If Not gM46V999.gfPreenchido(.TextMatrix(lI, ColPerLimite)) Then
                                    .TextMatrix(lI, ColPerLimite) = 0
                                End If
                                ' Plano Ipiranga - Inicio
                                If dValISLimite <> 0 Then
                                    .TextMatrix(lI, ColLimiteMaximo) = Format(dValISLimite, "standard")
                                Else
                                ' Plano Ipiranga - Fim
                                'Ficha 1000427
                                If gM46V999.gCodProduto > 987 Then
                                    If gM46V111.gCod_plano > 2000 Then    'planinho tem que arredondar
                                        .TextMatrix(lI, ColLimiteMaximo) = Format(Round((.TextMatrix(lFind, ColValLMG) * (.TextMatrix(lI, ColPerLimite) / 100) / 10)) * 10, "standard")
                                    Else
                                            .TextMatrix(lI, ColLimiteMaximo) = Format(((.TextMatrix(lFind, ColLimiteMaximo) * (.TextMatrix(lI, ColPerLimite) / 100) / 10)) * 10, "standard")
                                    End If
                                Else
                                    .TextMatrix(lI, ColLimiteMaximo) = Format(Round((.TextMatrix(lFind, ColValLMG) * (.TextMatrix(lI, ColPerLimite) / 100) / 10)) * 10, "standard")
                                End If
                                ' Plano Ipiranga - Inicio
                                End If
                                ' Plano Ipiranga - Fim
                                Exit For
                            End If
                        Next lFind
                    Else
                        If Not gM46V999.gfPreenchido(.TextMatrix(lI, ColPerLimite)) Then
                            .TextMatrix(lI, ColPerLimite) = "0"
                        End If
                        'Ficha 1000427
                        ' Plano Ipiranga - Inicio
                        If dValISLimite <> 0 Then
                            .TextMatrix(lI, ColLimiteMaximo) = Format(dValISLimite, "standard")
                        Else
                        ' Plano Ipiranga - Fim
                        If gM46V999.gCodProduto > 987 Then
                            If gM46V111.gCod_plano > 2000 Then    'planinho tem que arredondar
                                .TextMatrix(lI, ColLimiteMaximo) = Format(Round((lValorIS * .TextMatrix(lI, ColPerLimite) / 100) / 10) * 10, "standard")
                            Else
                                .TextMatrix(lI, ColLimiteMaximo) = Format(((lValorIS * .TextMatrix(lI, ColPerLimite) / 100) / 10) * 10, "standard")
                            End If
                        Else
                            .TextMatrix(lI, ColLimiteMaximo) = Format(Round((lValorIS * .TextMatrix(lI, ColPerLimite) / 100) / 10) * 10, "standard")
                        End If
                        ' Plano Ipiranga - Inicio
                        End If
                        ' Plano Ipiranga - Fim
                    End If
                End If
            End If
            If .TextMatrix(lI, ColValISAte) = "" Then .TextMatrix(lI, ColValISAte) = 0
            If CDbl(.TextMatrix(lI, ColValISAte)) > 0 And CDbl(.TextMatrix(lI, ColValISAte)) < CDbl(.TextMatrix(lI, ColLimiteMaximo)) Then
                .TextMatrix(lI, ColLimiteMaximo) = Format(.TextMatrix(lI, ColValISAte), "standard")
            End If
            
            If gM46V999.gCodProduto >= 1080 Then
                If gM46V999.CDblx(.TextMatrix(lI, ColValISmin)) > gM46V999.CDblx(.TextMatrix(lI, ColLimiteMaximo)) Then
                    .TextMatrix(lI, ColLimiteMaximo) = Format(.TextMatrix(lI, ColValISmin), "standard")
                End If
                .TextMatrix(lI, ColLimiteMaximo) = Format(AjustePontualValorMaximo3002(.TextMatrix(lI, ColCodCobert), gM46V999.CDblx(.TextMatrix(lI, ColLimiteMaximo)), gM46V999.CDblx(.TextMatrix(lI, ColValISmin))), "standard")
            End If
            
            If .TextMatrix(lI, ColCobObrigat) = "1" Then
                .TextMatrix(lI, ColValLMG) = .TextMatrix(lI, ColLimiteMaximo)
            End If

            ' Se for permitido alterar o Valor em Risco, pintar a celula de Branco.
            If lI <> 1 And mfBuscaCoberturasValorRiscoLMGAgravo(gM46V999.CDblx(.TextMatrix(lI, ColCodCobert))) >= 0# Then
                .CellBackColor = vbWhite
            End If

            If gM46V999.CDblx(.TextMatrix(lI, ColValRisco)) = 0 Or (.CellBackColor <> vbWhite And .CellBackColor <> &H80000018 And lI <> 1) Then
                .TextMatrix(lI, ColValRisco) = Format(.TextMatrix(lI, ColValLMG), "standard")
            Else
                .TextMatrix(lI, ColValRisco) = Format(.TextMatrix(lI, ColValRisco), "standard")
            End If
        Next lI
    End With

End Sub

Private Function mfRetornaLMG(pLinha As Long) As String
    Dim lRetornoFinal As Double
    Dim lObjstcPed As stcA46V702B
    Dim lobjstcPedLoc As stcA46V708B
    Dim lobjstcPedCobert As stcA46V704B
    Dim lretorno As String
    Dim lMensagem As String

    On Error GoTo Fim:
    
    lRetornoFinal = flxSelecionadas.TextMatrix(pLinha, ColLimiteMaximo)
    If gM46V999.gCodProduto >= 940 Then
        Set lObjstcPed = New stcA46V702B
        Set lObjstcPed.M46V999 = gM46V999
        Set lObjstcPed.objstcCtrlEmis = frmT46V101A.objstcPedido.objstcCtrlEmis

        Set lobjstcPedLoc = New stcA46V708B
        Set lobjstcPedLoc.M46V999 = gM46V999

        Set lobjstcPedCobert = New stcA46V704B
        Set lobjstcPedCobert.M46V999 = gM46V999


        lobjstcPedCobert.Cod_Cobert = flxSelecionadas.TextMatrix(pLinha, ColCodCobert)
        lobjstcPedLoc.Cod_Plano = frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex)

        If gM46V103.gfVerParametro(9, lObjstcPed, lobjstcPedLoc, lobjstcPedCobert, lretorno, lMensagem) = False Then
            Call gM46V999.gpGraLog(1, lMensagem)
            mfRetornaLMG = lRetornoFinal
            Exit Function
        End If

        If gM46V999.gfPreenchido(lretorno) Then
            lRetornoFinal = CDbl(lretorno)
            flxSelecionadas.TextMatrix(pLinha, ColLimiteMaximo) = Format(lRetornoFinal, "standard")
        End If

        If gM46V999.gCodProduto >= 1080 Then
        Dim lValLimMaxCobert As Currency
        
        Set lobjstcPedLoc.ObjPai = lObjstcPed
        Set lobjstcPedCobert.ObjPai = lobjstcPedLoc
        
        lObjstcPed.Cod_Produto = gM46V999.gCodProduto
        lObjstcPed.Cod_Estip = lobjstcPedLoc.Cod_Plano
        lobjstcPedCobert.Cod_Ramo = gM46V999.gCod_Ramo
        
        lValLimMaxCobert = lobjstcPedCobert.gVal_Lim_Max_Cobert
        If lValLimMaxCobert > 0 Then
        mfRetornaLMG = Format(lValLimMaxCobert, "standard")
        Else
            mfRetornaLMG = lRetornoFinal
        End If
        Else
            If lRetornoFinal = 0 Then
                lRetornoFinal = "999999999999999"
            End If
            mfRetornaLMG = Format(lRetornoFinal, "standard")
        End If

    Else
        mfRetornaLMG = Format("999999999999999", "standard")
    End If
    
Fim:
    Set lObjstcPed = Nothing
    Set lobjstcPedLoc = Nothing
    Set lobjstcPedCobert = Nothing
End Function

Private Function mfRetornaLMGMin(pLinha As Long) As String
    Dim lRetornoFinal As Double
    Dim lObjstcPed As stcA46V702B
    Dim lobjstcPedLoc As stcA46V708B
    Dim lobjstcPedCobert As stcA46V704B
    Dim lretorno As String
    Dim lMensagem As String

    lRetornoFinal = 0
    
    Set lObjstcPed = New stcA46V702B
    Set lObjstcPed.M46V999 = gM46V999
    Set lObjstcPed.objstcCtrlEmis = frmT46V101A.objstcPedido.objstcCtrlEmis
    
    Set lobjstcPedLoc = New stcA46V708B
    Set lobjstcPedLoc.M46V999 = gM46V999
    
    Set lobjstcPedCobert = New stcA46V704B
    Set lobjstcPedCobert.M46V999 = gM46V999
    
    lobjstcPedCobert.Cod_Cobert = flxSelecionadas.TextMatrix(pLinha, ColCodCobert)
    lobjstcPedLoc.Cod_Plano = frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex)
    
    If gM46V103.gfVerParametro(9, lObjstcPed, lobjstcPedLoc, lobjstcPedCobert, lretorno, lMensagem) = False Then
        Call gM46V999.gpGraLog(1, lMensagem)
        mfRetornaLMGMin = lRetornoFinal
        Exit Function
    End If
    
    If gM46V999.gfPreenchido(lretorno) Then
        lRetornoFinal = CDbl(lretorno)
        flxSelecionadas.TextMatrix(pLinha, ColLimiteMaximo) = Format(lRetornoFinal, "standard")
    End If
    
    Dim lValLimMinCobert As Currency
    
    Set lobjstcPedLoc.ObjPai = lObjstcPed
    Set lobjstcPedCobert.ObjPai = lobjstcPedLoc
    
    lObjstcPed.Cod_Produto = gM46V999.gCodProduto
    lObjstcPed.Cod_Estip = lobjstcPedLoc.Cod_Plano
    lobjstcPedCobert.Cod_Ramo = gM46V999.gCod_Ramo
    
    lValLimMinCobert = lobjstcPedCobert.gVal_Lim_Min_Cobert
    lRetornoFinal = Format(lValLimMinCobert, "standard")
    
    Set lObjstcPed = Nothing
    Set lobjstcPedLoc = Nothing
    Set lobjstcPedCobert = Nothing
    
    mfRetornaLMGMin = lRetornoFinal
End Function

Private Function mfBuscaPerIndenitarioDefault(ByVal pCodCobert As Integer) As Integer
    Dim lbdTbm_Coef_PI As ADODB.Connection    'Arquivo: P00Multi.
    Dim lMensagem As String         'Mensagem.
    Dim lrsTbm_Coef_PI As clsYasRecordSet     'Registro : Tbm_Coef_PI.
    Dim lSelect As String
    Dim Cod_Modalidade As Integer

    Cod_Modalidade = frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex)

    mfBuscaPerIndenitarioDefault = 12

    If Me.Tag = "Desabilitado" Then
        Exit Function
    End If

    If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTbm_Coef_PI, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Function
    End If

    '1 - Procura pela Cobertura na Tabela de PI Default


    lSelect = " Select *  from Tbm_Per_Ind_Padrao "
    lSelect = lSelect & " Where Cod_Ramo = " & gM46V999.gCod_Ramo
    lSelect = lSelect & " and Cod_Modalidade = " & Cod_Modalidade
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
        mfBuscaPerIndenitarioDefault = lrsTbm_Coef_PI("Per_Indenit_Def")
    Else
        mfBuscaPerIndenitarioDefault = 12    ' é o padrão
    End If

    Call gM46V999.gpFecha2(lbdTbm_Coef_PI)

End Function

' Tarifa 1056 e 1057 - Inicio
Private Function mfBuscaCoberturasValorRiscoLMGAgravo(ByVal pCodCobert As Integer) As Double
    Dim lbdTbm_Coef_PI As ADODB.Connection    'Arquivo: P00Multi.
    Dim lMensagem As String         'Mensagem.
    Dim lrsTbm_Coef_PI As clsYasRecordSet     'Registro : Tbm_Coef_PI.
    Dim lSelect As String

    mfBuscaCoberturasValorRiscoLMGAgravo = -1

    If Me.Tag = "Desabilitado" Then
        Exit Function
    End If

    If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTbm_Coef_PI, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Function
    End If

    '1 - Procura pela Cobertura na Tabela de PI Default


    lSelect = " Select *  from Tbm_VR_LMG_Agrav "
    lSelect = lSelect & " Where Cod_Ramo = " & gM46V999.gCod_Ramo
    lSelect = lSelect & " and Cod_Cobert = " & pCodCobert
    lSelect = lSelect & " and Cod_Tarif = " & gM46V999.gCodProduto

    ' REMOVER ISSO - Incluido apenas para teste debug
    mfBuscaCoberturasValorRiscoLMGAgravo = -1
    Exit Function
    ' REMOVER ISSO - Incluido apenas para teste debug

    If gM46V999.gfObtRegistro(lbdTbm_Coef_PI, lSelect, lrsTbm_Coef_PI, lMensagem) = False Then
        Call gM46V999.gpFecha2(lbdTbm_Coef_PI)
        Call gM46V999.gpGraLog(2, lMensagem)
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    '2 - Se encontrar permite cadastro de Período Indenitário
    If Not lrsTbm_Coef_PI.EOF Then
        If lrsTbm_Coef_PI.RecordCount <> 1 Then
            mfBuscaCoberturasValorRiscoLMGAgravo = 999
        Else
            mfBuscaCoberturasValorRiscoLMGAgravo = lrsTbm_Coef_PI("Percent_Agravo")
        End If
    Else
        mfBuscaCoberturasValorRiscoLMGAgravo = -1    ' Não encontrou - não permite (-1 é False no VB)
    End If

    Call gM46V999.gpFecha2(lbdTbm_Coef_PI)

End Function
' Tarifa 1056 e 1057 - Fim

Public Function ValidaPI()
'CR23794 - Syas Kit - Erro PI
    Dim lI As Integer
    If gPiDefault Then
        With flxSelecionadas
            For lI = 1 To .Rows - 1
                If .TextMatrix(lI, ColPermitePInd) = "True" Then
                    .TextMatrix(lI, ColPInd) = mfBuscaPerIndenitarioDefault(.TextMatrix(lI, ColCodCobert))
                End If
            Next lI
            gPiDefault = False
        End With
    End If
End Function

Public Sub mpRecalcularLMGLC()
    Dim lValor As Double
    Dim lI As Integer

    lValor = 0
    For lI = 1 To flxSelecionadas.Rows - 1
        If flxSelecionadas.TextMatrix(lI, ColCodGrupoCobert) = "3" And val(flxSelecionadas.TextMatrix(lI, ColValLMG)) > 0 Then
            lValor = lValor + flxSelecionadas.TextMatrix(lI, ColValLMG)
        End If
    Next lI
    frmT46V101A.txtLucros_Cessantes.Text = Format(lValor, "standard")
    
    ' Para planinhos deve repetir o valor dos lucros cessantes
    If frmT46V101A.objstcPedido.getPlano.EhPlaninho = True Then
        frmT46V101A.txtValoRisco_LC.Text = Format(lValor, "standard")
    End If

    If val(frmT46V101A.txtValoRisco_LC.Text) > 0 And val(frmT46V101A.txtValoRisco_LC.Text) > 0 Then
        frmT46V101A.txtCoefISVR_LC.Text = frmT46V101A.txtLucros_Cessantes.Text / frmT46V101A.txtValoRisco_LC.Text
    Else
        frmT46V101A.txtCoefISVR_LC.Text = Format("0", "standard")
    End If
End Sub

Private Function FormataDesconto(ByVal pDesconto As Double) As Integer
    On Error GoTo Erro
    
    If pDesconto > 0 Then
        FormataDesconto = CInt((1# - pDesconto) * 100)
    Else
        FormataDesconto = 0
    End If
    Exit Function
Erro:
    FormataDesconto = 0
    MsgBox "Erro ao calcular o desconto!", vbInformation
End Function

Public Sub FormataDanosEletricos()
    Dim lDescricao As String
    Dim i As Integer
    
    With flxSelecionadas
        For i = 1 To .Rows - 1
            If .TextMatrix(i, ColCodCobert) = 1103 And gM46V999.gCod_Ramo = e_CodRamo_Condominio Then
                If InStr(.TextMatrix(i, ColNomCobert), " COM ELEVADOR") > 0 Then
                    lDescricao = RTrim(Left(.TextMatrix(i, ColNomCobert), InStr(.TextMatrix(i, ColNomCobert), " COM ELEVADOR")))
                Else
                    If InStr(.TextMatrix(i, ColNomCobert), " SEM ELEVADOR") > 0 Then
                        lDescricao = RTrim(Left(.TextMatrix(i, ColNomCobert), InStr(.TextMatrix(i, ColNomCobert), " SEM ELEVADOR")))
                    Else
                        lDescricao = .TextMatrix(i, ColNomCobert)
                    End If
                End If
                
                Select Case frmT46V101A.cboElevador.ItemData(frmT46V101A.cboElevador.ListIndex)
                Case 0
                    .TextMatrix(i, ColNomCobert) = lDescricao
                Case 1
                    .TextMatrix(i, ColNomCobert) = lDescricao & " COM ELEVADOR"
                Case 2
                    .TextMatrix(i, ColNomCobert) = lDescricao & " SEM ELEVADOR"
                End Select
            End If
        Next
    End With

End Sub

Private Function AjustePontualValorMaximo3002(pCod_Cobert As String, pValorMaximo As Currency, pValorMinimo As Currency) As Currency
    
    AjustePontualValorMaximo3002 = pValorMaximo
    
    If gM46V999.gCod_Ramo = e_CodRamo_Condominio And gM46V999.gCodProduto >= 1080 Then
        
        If pCod_Cobert = "3002" Then
        
            With flxSelecionadas
                    
                If .TextMatrix(1, ColValLMG) > 0 And _
                (.TextMatrix(1, ColValLMG) * 0.1) > pValorMinimo Then
                    AjustePontualValorMaximo3002 = (.TextMatrix(1, ColValLMG) * 0.1)
                End If
                
                Dim i As Integer
                
                pValorMaximo = AjustePontualValorMaximo3002
                
                For i = 1 To .Rows - 1
                    If .TextMatrix(i, ColCodCobert) = 1691 Then
                        If .TextMatrix(i, ColValLMG) > 0 And _
                           CCur(.TextMatrix(i, ColValLMG)) < pValorMaximo And _
                           .TextMatrix(i, ColValLMG) > pValorMinimo Then
                           
                            AjustePontualValorMaximo3002 = .TextMatrix(i, ColValLMG)
                        End If
                        Exit For
                    End If
                    
                Next
                
            End With
        
        End If
        
    End If
    
End Function

''20160909 - Bug de Franquia Dias/Valor
Public Sub mpAtualizaFranquiaGeral()
    Dim lI As Integer
    For lI = 1 To flxSelecionadas.Rows - 1
        flxSelecionadas.TextMatrix(lI, ColAlterouFranquia) = ""
        Call mpAtualizaFranquiaMin(lI)
    Next
End Sub

