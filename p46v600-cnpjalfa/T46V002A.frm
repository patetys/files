VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT46V002A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Renovação"
   ClientHeight    =   2730
   ClientLeft      =   1275
   ClientTop       =   2400
   ClientWidth     =   8595
   ControlBox      =   0   'False
   Icon            =   "T46V002A.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   8595
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel panRenYasuda 
      Height          =   2100
      Left            =   2430
      TabIndex        =   3
      Top             =   120
      Width           =   6045
      _Version        =   65536
      _ExtentX        =   10663
      _ExtentY        =   3704
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      Begin VB.CheckBox chkCorrecaoMonetariaIPC 
         BackColor       =   &H80000004&
         Caption         =   "Corrigir valores pelo IPC"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   720
         TabIndex        =   10
         Top             =   1710
         Visible         =   0   'False
         Width           =   2025
      End
      Begin VB.TextBox txtRenItem 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   3795
         MaxLength       =   4
         TabIndex        =   9
         ToolTipText     =   "Número do item - renovação Yasuda."
         Top             =   1365
         Width           =   720
      End
      Begin VB.OptionButton optItem 
         Caption         =   "Todos os itens"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   720
         TabIndex        =   7
         Top             =   1365
         Value           =   -1  'True
         Width           =   1500
      End
      Begin VB.OptionButton optItem 
         Caption         =   "Apenas item "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   1
         Left            =   2520
         TabIndex        =   8
         Top             =   1365
         Width           =   1500
      End
      Begin VB.TextBox txtRenApolice 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1500
         MaxLength       =   10
         TabIndex        =   5
         ToolTipText     =   "Número da apólice - Yasuda"
         Top             =   915
         Width           =   1200
      End
      Begin Threed.SSCommand cmdRenovacao 
         Height          =   285
         Left            =   2760
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Busca de apólice."
         Top             =   915
         Width           =   300
         _Version        =   65536
         _ExtentX        =   529
         _ExtentY        =   503
         _StockProps     =   78
         Picture         =   "T46V002A.frx":000C
      End
      Begin VB.Label lblX 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Apólice: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   690
         TabIndex        =   4
         Top             =   915
         Width           =   795
      End
   End
   Begin Threed.SSPanel panRenMaritima 
      Height          =   2100
      Left            =   2430
      TabIndex        =   22
      Top             =   120
      Width           =   6045
      _Version        =   65536
      _ExtentX        =   10663
      _ExtentY        =   3704
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      Begin VB.TextBox txtApoliceMaritima 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1530
         MaxLength       =   6
         TabIndex        =   28
         ToolTipText     =   "Número da apólice - Yasuda"
         Top             =   960
         Width           =   690
      End
      Begin VB.ComboBox cboSucursal 
         ForeColor       =   &H00800000&
         Height          =   315
         ItemData        =   "T46V002A.frx":0326
         Left            =   1530
         List            =   "T46V002A.frx":0328
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   180
         Width           =   2715
      End
      Begin VB.ComboBox cboRamoMaritima 
         ForeColor       =   &H00800000&
         Height          =   315
         ItemData        =   "T46V002A.frx":032A
         Left            =   1530
         List            =   "T46V002A.frx":032C
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   570
         Width           =   2715
      End
      Begin VB.TextBox txtRenItemMaritima 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   3795
         MaxLength       =   4
         TabIndex        =   31
         ToolTipText     =   "Número do item - renovação Yasuda."
         Top             =   1365
         Width           =   720
      End
      Begin VB.OptionButton optItemMaritima 
         Caption         =   "Apenas item "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   1
         Left            =   2520
         TabIndex        =   30
         Top             =   1365
         Width           =   1170
      End
      Begin VB.OptionButton optItemMaritima 
         Caption         =   "Todos os itens"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   720
         TabIndex        =   29
         Top             =   1365
         Value           =   -1  'True
         Width           =   1500
      End
      Begin VB.Label lblX 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Apólice: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   4
         Left            =   720
         TabIndex        =   27
         Top             =   960
         Width           =   795
      End
      Begin VB.Label lblSucursal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sucursal: "
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   720
         TabIndex        =   23
         Top             =   180
         Width           =   795
      End
      Begin VB.Label lblRamo 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ramo: "
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   720
         TabIndex        =   25
         Top             =   570
         Width           =   795
      End
   End
   Begin VB.OptionButton optTipEmissao 
      Caption         =   "Renovação de seguro &Marítima"
      ForeColor       =   &H00800000&
      Height          =   495
      Index           =   3
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Renovação Yasuda."
      Top             =   720
      Width           =   1995
   End
   Begin VB.OptionButton optTipEmissao 
      Caption         =   "Renovação de seguro &Yasuda"
      ForeColor       =   &H00800000&
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Renovação Yasuda."
      Top             =   120
      Width           =   1995
   End
   Begin VB.OptionButton optTipEmissao 
      Caption         =   "Renovação de &seguro de outra seguradora (congênere)"
      ForeColor       =   &H00800000&
      Height          =   855
      Index           =   2
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Renovação de congênere."
      Top             =   1320
      Width           =   2175
   End
   Begin Threed.SSCommand cmdOk 
      Height          =   300
      Left            =   3195
      TabIndex        =   32
      ToolTipText     =   "Ok."
      Top             =   2250
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
      Left            =   4395
      TabIndex        =   33
      ToolTipText     =   "Cancelar."
      Top             =   2250
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Cancelar"
      ForeColor       =   8388608
   End
   Begin Threed.SSPanel panRenCongenere 
      Height          =   1740
      Left            =   2430
      TabIndex        =   11
      Top             =   120
      Width           =   6045
      _Version        =   65536
      _ExtentX        =   10663
      _ExtentY        =   3069
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      Begin VB.TextBox txtSegItem 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3100
         MaxLength       =   5
         TabIndex        =   17
         ToolTipText     =   "Item da apólice de congênere."
         Top             =   900
         Width           =   700
      End
      Begin VB.TextBox txtSegApolice 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3100
         MaxLength       =   15
         TabIndex        =   15
         ToolTipText     =   "Número de apólice de congênere."
         Top             =   500
         Width           =   1700
      End
      Begin VB.ComboBox cboSeguradora 
         ForeColor       =   &H00800000&
         Height          =   315
         ItemData        =   "T46V002A.frx":032E
         Left            =   1200
         List            =   "T46V002A.frx":0330
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   13
         ToolTipText     =   "Nome da seguradora congênere."
         Top             =   100
         Width           =   4665
      End
      Begin Threed.SSPanel panX 
         Height          =   300
         Index           =   0
         Left            =   3100
         TabIndex        =   19
         Top             =   1300
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2469
         _ExtentY        =   529
         _StockProps     =   15
         ForeColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   0
         BevelOuter      =   0
         BevelInner      =   1
         Alignment       =   6
         Begin VB.OptionButton optSinistro 
            Caption         =   "&Não"
            ForeColor       =   &H00800000&
            Height          =   200
            Index           =   1
            Left            =   700
            TabIndex        =   21
            ToolTipText     =   "NÃO tem cobertura de despesas extraordinárias."
            Top             =   40
            Value           =   -1  'True
            Width           =   600
         End
         Begin VB.OptionButton optSinistro 
            Caption         =   "&Sim"
            ForeColor       =   &H00800000&
            Height          =   200
            Index           =   0
            Left            =   100
            TabIndex        =   20
            ToolTipText     =   "SIM, tem cobertura de despesas extraordinárias."
            Top             =   40
            Width           =   600
         End
      End
      Begin VB.Label lblX 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Item: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   3
         Left            =   100
         TabIndex        =   16
         Top             =   900
         Width           =   3000
      End
      Begin VB.Label lblX 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Apólice: "
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   2
         Left            =   100
         TabIndex        =   14
         Top             =   500
         Width           =   3000
      End
      Begin VB.Label lblX 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Seguradora: "
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   1
         Left            =   100
         TabIndex        =   12
         Top             =   100
         Width           =   1100
      End
      Begin VB.Label lblX 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ocorreu sinistro? "
         ForeColor       =   &H00800000&
         Height          =   300
         Index           =   5
         Left            =   100
         TabIndex        =   18
         Top             =   1300
         Width           =   3000
      End
   End
End
Attribute VB_Name = "frmT46V002A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mNosNumero As String  'Nosso número informado pelos usuário de digitação.
Private Function mfProRenCongenere(Optional ByRef pNosNumero As String, Optional ByRef pMensagem As String) As Boolean
'Função: processa opção de renovação de congênere.

    Dim clsA99V001 As clsA99V001  'Classe de cálculo de dígitos.
    Dim lCodIdeApolice As String          'Código de identificação da apólice.
    Dim lCodProduto As Integer      'Código do produto.
    Dim lCodProtocolo As Long           'Código do protocolo.
    Dim lCodSequencia As Long           'Código da seqüência.
    Dim lCodSeguradora As Integer         'Código da seguradora.
    Dim lDescTarifa As String       'Descrição da tarifa.
    Dim lDigito As String   'Dígito.
    Dim lIndSinistro As Byte          'Indicador de sinistro.
    Dim lMensagem As String     'Mensagem.
    Dim lNosNumero As String      'Nosso número.
    Dim lNumApolice As String       'Número da apólice.
    Dim lNumItem As String    'Número do item.
    Dim lNumPedido As String      'Número do pedido.
    Dim lPosicao As Byte      'Utilizado com InStr.
    Dim lVigDatIni As Double      'Data de vigência: início.
    Dim lVigDatTer As Double      'Data de vigência: término.
    Dim lI As Integer    'Utilizafo no for...next.
    Dim lObjSegur As clsA46V716A    'Objeto da classe clsA46V716A (TAB_PED_LOG_SEGURO).
    Dim lObjstcSegur As stcA46V716B   'Objeto da classe de estrutura stcA46V716B (TAB_PED_LOC_SEGURO).
    Dim bdP00MPED As ADODB.Connection    'Conexão.

    On Error GoTo ControleErro

    mfProRenCongenere = False

    '1000347 - Inclusão de informações da congênere em renovações --Inicio
    'Posiciona seguradora.
    If cboSeguradora.ListIndex = -1 Then
        Call gM46V999.gpGraLog(1, "FGE0094 - Seguradora inválida (não selecionada).")
        If cboSeguradora.Visible = True And cboSeguradora.Enabled = True Then
            cboSeguradora.SetFocus
        End If
        Exit Function
    End If
    lCodSeguradora = cboSeguradora.ItemData(cboSeguradora.ListIndex)

    'Posiciona indicador de sinistro.
    If optSinistro(0) = True Then
        lIndSinistro = 1
    Else
        If optSinistro(1) = True Then
            lIndSinistro = 2
        Else
            lIndSinistro = 1
        End If
    End If

    If gM46V999.gfPreenchido(txtSegApolice.Text) Then
        lNumApolice = txtSegApolice.Text
    Else
        lNumApolice = " "
    End If
    If gM46V999.gfPreenchido(txtSegItem.Text) Then
        lNumItem = txtSegItem.Text
    Else
        lNumItem = " "
    End If
    '1000347 - Inclusão de informações da congênere em renovações --Fim

    '1. Calcula nosso número.

    If gM46V999.gfPreenchido(pNosNumero) Then
        lNosNumero = pNosNumero
    Else
        If gM46V999.gfCalNosNumero(lNosNumero, lMensagem) = False Then
            Call gM46V999.gpGraLog(1, lMensagem)
            Exit Function
        End If
        pNosNumero = lNosNumero
    End If

    mNosNumero = pNosNumero

    '2. Abre base de dados.
    If gM46V999.gfAbrBasDados("P00MPED", gM46V999.gTipP00MPED, gM46V999.gEndP00MPED, bdP00MPED, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        End
    End If

    '5. Calcula número do pedido.
    If gM46V999.gfCalNumPedido(lNumPedido, lMensagem, bdP00MPED) = False Then
        Exit Function
    End If
    gM46V999.gPrefixo = "Cotação " & lNumPedido & ". "

    '6. Posiciona período de vigência.
    lVigDatIni = CDbl(Format$(Now, "yyyymmdd"))
    lVigDatTer = lVigDatIni + 10000

    '7. Busca produto que atenda à vigência.
    If gM46V999.gfBusProduto(lVigDatIni, lCodProduto, lCodProtocolo, lCodSequencia, lDescTarifa, lMensagem, 0) = False Then
        Call gM46V999.gpGraLog(1, lMensagem)
        Exit Function
    End If

    '8. Grava registro de dados comuns.
    If gM46V001.gfGraNovPedido(bdP00MPED, lNosNumero, lNumPedido, lVigDatIni, lVigDatTer, lCodProduto, lCodProtocolo, _
                               lCodSequencia, lMensagem, 0, 0, 0, Nothing, "") = False Then
        Call gM46V999.gpGraLog(1, lMensagem)
        Exit Function
    End If

    '1000347 - Inclusão de informações da congênere em renovações
    '09. Grava o item
    If gM46V999.gfGraItem(bdP00MPED, lNosNumero, lNumPedido, "0001", lMensagem, _
                          e_TipEmissao_RenCongenere, , lCodSeguradora, lNumApolice, lNumItem, lIndSinistro) = False Then
        Call gM46V999.gpGraLog(1, lMensagem)
        Exit Function
    End If

    '10. Cria arquivo de comunicação.
    If gM46V111.gfGraRegComunicacao(lNosNumero, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Function
    End If

    mfProRenCongenere = True
    Exit Function
ControleErro:
    Call gM46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gM46V999.gpGraLog(0, "mfProRenCongenere ERRO - " & Err & " " & Err.Description)
End Function
Private Function mfProRenMaritima() As Boolean
    Dim lMensagem   As String                 'Mensagem.
    Dim lNumApolice As String

    On Error GoTo ControleErro

    mfProRenMaritima = False

    If gM46V999.gfPreenchido(txtApoliceMaritima.Text) Then
        lNumApolice = Right("0000000000" & Right("00" & cboSucursal.ItemData(cboSucursal.ListIndex), 2) & Right("00" & cboRamoMaritima.ItemData(cboRamoMaritima.ListIndex), 2) & Right("000000" & txtApoliceMaritima.Text, 6), 10)
    Else
        lNumApolice = " "
    End If
    
    If optItemMaritima(1).Value = True Then    'Apenas um item.
        If Not gM46V999.gfPreenchido(txtRenItemMaritima.Text) Then
            Call gM46V999.gpGraLog(1, "FGE0070 - Número do item da apólice inválido (renovação Marítima - não " & _
                                      "preenchido).")
            If txtRenItemMaritima.Visible = True And txtRenItemMaritima.Enabled = True Then
                txtRenItemMaritima.SetFocus
            End If
            Exit Function
        End If
        If Not IsNumeric(txtRenItemMaritima.Text) Then
            Call gM46V999.gpGraLog(1, "FGE0070 - Número do item da apólice inválido (renovação Marítima - não " & _
                                      "numérico).")
            If txtRenItemMaritima.Visible = True And txtRenItemMaritima.Enabled = True Then
                txtRenItemMaritima.SetFocus
            End If
            Exit Function
        End If
        If Len(Trim(txtRenItemMaritima.Text)) < 4 Then
            txtRenItemMaritima.Text = Trim$(txtRenItemMaritima.Text)
            txtRenItemMaritima.Text = String(4 - Len(txtRenItemMaritima.Text), "0") & txtRenItemMaritima.Text
        End If
    End If
    
    lMensagem = ""
    '3. Grava registros de cotação/proposta para renovação Yasuda.
    If gM46V001.gfGraRenYasuda(lNumApolice, txtRenItemMaritima.Text, mNosNumero, _
                               lMensagem, lMensagem, "") = False Then
        If gM46V999.gfPreenchido(lMensagem) Then
            Call gM46V999.gpGraLog(1, lMensagem)
        End If
        If txtRenApolice.Visible = True And txtRenApolice.Enabled = True Then
            txtRenApolice.SetFocus
        End If
        Exit Function
    End If

    mfProRenMaritima = True
    Exit Function
ControleErro:
    Call gM46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gM46V999.gpGraLog(0, "mfProRenMaritima ERRO - " & Err & " " & Err.Description)
End Function
Private Function mfProRenYasuda() As Boolean
'Função: processa opção de renovação de seguro Yasuda.

    Dim lMensagem As String    'Mensagem.

    mfProRenYasuda = False
    On Error GoTo ControleErro

    '1. Consiste número de apólice anterior.
    If Not gM46V999.gfPreenchido(txtRenApolice.Text) Then
        Call gM46V999.gpGraLog(1, "FGE0069 - Número da apólice inválido (renovação Yasuda - não preenchido).")
        If txtRenApolice.Visible = True And txtRenApolice.Enabled = True Then
            txtRenApolice.SetFocus
        End If
        Exit Function
    End If
    If Not IsNumeric(txtRenApolice.Text) Then
        Call gM46V999.gpGraLog(1, "FGE0069 - Número da apólice inválido (renovação Yasuda - não numérico).")
        If txtRenApolice.Visible = True And txtRenApolice.Enabled = True Then
            txtRenApolice.SetFocus
        End If
        Exit Function
    End If
    If Len(Trim(txtRenApolice.Text)) < 10 Then
        txtRenApolice.Text = Trim$(txtRenApolice.Text)
        txtRenApolice.Text = String(10 - Len(txtRenApolice.Text), "0") & txtRenApolice.Text
    End If

    '2. Consiste número do item da apólice.
    If optItem(1) = True Then   'Apenas um item.
        If Not gM46V999.gfPreenchido(txtRenItem.Text) Then
            Call gM46V999.gpGraLog(1, "FGE0070 - Número do item da apólice inválido (renovação Yasuda - não " & _
                                      "preenchido).")
            If txtRenItem.Visible = True And txtRenItem.Enabled = True Then
                txtRenItem.SetFocus
            End If
            Exit Function
        End If
        If Not IsNumeric(txtRenItem.Text) Then
            Call gM46V999.gpGraLog(1, "FGE0070 - Número do item da apólice inválido (renovação Yasuda - não " & _
                                      "numérico).")
            If txtRenItem.Visible = True And txtRenItem.Enabled = True Then
                txtRenItem.SetFocus
            End If
            Exit Function
        End If
        If Len(Trim(txtRenItem.Text)) < 4 Then
            txtRenItem.Text = Trim$(txtRenItem.Text)
            txtRenItem.Text = String(4 - Len(txtRenItem.Text), "0") & txtRenItem.Text
        End If
    End If
    lMensagem = ""
    lMensagem = InputBox("Digite a tarifa que será usada na renovação (em branco usará a atual): ", Me.Caption)
    
    Dim lMsg_Erro As String
    Dim lTarifaAtual_Impar As Integer
    
    lTarifaAtual_Impar = gM46V999.gfTarifaAtual_Impar(lMsg_Erro)
    If gM46V999.gfPreenchido(lMsg_Erro) Then
        Call gM46V999.gpGraLog(1, "FGE0070 - " & lMsg_Erro)
        Exit Function
    End If
    
    If lMensagem <> Empty Then
        If Not IsNumeric(lMensagem) Then
            Call gM46V999.gpGraLog(1, "FGE0070 - Tarifa inválida.")
            Exit Function
        Else
            If lMensagem > lTarifaAtual_Impar Then
                Call gM46V999.gpGraLog(1, "FGE0070 - Tarifa inválida.")
                Exit Function
            End If
        End If
    End If
    
    '3. Grava registros de cotação/proposta para renovação Yasuda.
'    gM46V999.gCorrecaoMonetariaIPC = (chkCorrecaoMonetariaIPC.Value = Checked)
    If gM46V001.gfGraRenYasuda(txtRenApolice.Text, txtRenItem.Text, mNosNumero, _
                               lMensagem, lMensagem, "") = False Then
        If gM46V999.gfPreenchido(lMensagem) Then
            Call gM46V999.gpGraLog(1, lMensagem)
        End If
        If txtRenApolice.Visible = True And txtRenApolice.Enabled = True Then
            txtRenApolice.SetFocus
        End If
        Exit Function
    End If
    'Alerta do LMI
    If gM46V999.gfPreenchido(lMensagem) Then
        Call gM46V999.gpGraLog(1, lMensagem)
    End If
    mfProRenYasuda = True
    Exit Function
ControleErro:
    Call gM46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gM46V999.gpGraLog(0, "mpProRenYasuda ERRO - " & Err & " " & Err.Description)
End Function
Public Sub mpInstanciarForm(pNosNumero As String)
    mNosNumero = pNosNumero
    Me.Caption = "Cotação - Renovação"
    Me.Show 1
    pNosNumero = mNosNumero
End Sub

'1000347 - Inclusão de informações da congênere em renovações
Private Sub cboSeguradora_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        'KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cmdCancelar_Click()
    gM46V999.gFormCancelado = True
    Unload Me
End Sub
Private Sub cmdOk_Click()
    Screen.MousePointer = vbHourglass

    On Error GoTo ControleErro

    '1. Consiste opção.
    If optTipEmissao(1) = False And optTipEmissao(2) = False And optTipEmissao(3) = False Then
        Call gM46V999.gpGraLog(1, "FNP0016 - Opção não selecionada.")
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    '2. Processa opção.
    If optTipEmissao(1) = True Then
        If mfProRenYasuda = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    Else
        If optTipEmissao(2) = True Then
            If mfProRenCongenere = False Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        Else
            If optTipEmissao(3) = True Then
                If mfProRenMaritima = False Then
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            End If
        End If
    End If
    
    '3. Retorna ao primeiro módulo.
    txtRenApolice.Text = ""
    txtRenItem.Text = ""

    Screen.MousePointer = vbDefault
    On Error Resume Next
    Me.Hide
    On Error GoTo 0
    Exit Sub
ControleErro:
    Call gM46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gM46V999.gpGraLog(0, "T46V002A.cmdOk_Click ERRO - " & Err & " " & Err.Description)
End Sub

Private Sub cmdRenovacao_Click()
    Dim lNum_Apolice As String  'Número da apólice
    Dim lNum_Item As String    'Número do item

    On Error GoTo ControleErro
    Call frmT46V003A.mpInstanciarForm(lNum_Apolice, lNum_Item)
    txtRenApolice.Text = lNum_Apolice
    If cmdOk.Visible = True And cmdOk.Enabled = True Then
        cmdOk.SetFocus
    End If
    Exit Sub
ControleErro:
    Call gM46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gM46V999.gpGraLog(0, "T46V002.cmdRenovacao_click ERRO - " & Err & " " & Err.Description)
End Sub

Private Sub Form_Activate()
    frmT46V001A.panForm = Mid$(Me.Name, 4)
    If gM46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then
        If gM46V999.gDesabFormsTransmi = True Then
            DesabilitaControles Me
            cmdOk.Enabled = True
            cmdCancelar.Enabled = True
        End If
    End If
End Sub
Private Sub Form_Load()
    Call gM46V999.gpCenForm(Me)  'Centraliza form.

    '1. Posiciona opção de renovação Yasuda.
    optTipEmissao(1).Value = True

    '1000347 - Inclusão de informações da congênere em renovações
    txtSegApolice.Text = ""
    txtSegItem.Text = ""
    Call Carregar_CBO_SucursalMaritima
    Call Carregar_CBO_RamoMaritima
End Sub
Private Sub Form_Unload(Cancel As Integer)
    mNosNumero = ""
    Set frmT46V002A = Nothing
End Sub
Private Sub optItem_Click(index As Integer)
    Select Case index
    Case 0      'Todos os itens.
        txtRenItem.Text = ""
        txtRenItem.Enabled = False
        txtRenItem.BackColor = &H8000000F
    Case 1      'Apenas o item.
        txtRenItem.Enabled = True
        txtRenItem.BackColor = vbWhite
    End Select
End Sub
Private Sub optItem_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub optItemMaritima_Click(index As Integer)
    Select Case index
    Case 0      'Todos os itens.
        txtRenItemMaritima.Text = ""
        txtRenItemMaritima.Enabled = False
        txtRenItemMaritima.BackColor = &H8000000F
    Case 1      'Apenas o item.
        txtRenItemMaritima.Enabled = True
        txtRenItemMaritima.BackColor = vbWhite
    End Select
End Sub
Private Sub optItemMaritima_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub optTipEmissao_Click(index As Integer)

    Dim bdP0044700 As ADODB.Connection      'Arquivo - P0044700.
    Dim lMensagem As String                 'Mensagem.
    Dim rsP0044700 As clsYasRecordSet       'Registro - P0044700.
    Dim lSelect As String

    Screen.MousePointer = vbHourglass
    On Error GoTo ControleErro

    Select Case index
    Case 1      'Renovação de seguro Yasuda.
        '1000347 - Inclusão de informações da congênere em renovações
        panRenCongenere.Visible = False
        panRenMaritima.Visible = False
        panRenYasuda.Visible = True
        
    Case 3      'Renovação de seguro Maritima.
        panRenCongenere.Visible = False
        panRenYasuda.Visible = False
        panRenMaritima.Visible = True
        
    Case 2      'Renovação de seguro de outras seguradoras.
        '1000347 - Inclusão de informações da congênere em renovações
        If cboSeguradora.ListCount = 0 Then
            If gM46V999.gfAbrBasDados("P0044700", gM46V999.gTipP0044700, gM46V999.gEndP0044700, bdP0044700, lMensagem) = False _
               Then
                Call gM46V999.gpGraLog(2, lMensagem)
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            lSelect = "SELECT CSeguradora, Seguradora FROM YTAB_SEGURADORA"
            If gM46V999.gfObtRegistro(bdP0044700, lSelect, rsP0044700, lMensagem) = False Then
                Call gM46V999.gpFecha2(bdP0044700)
                Call gM46V999.gpGraLog(2, lMensagem)
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            If rsP0044700.EOF = False Then
                With cboSeguradora
                    While Not rsP0044700.EOF
                        Select Case val(rsP0044700!CSeguradora)
                        Case 641
                            'Yasuda - não incluir.
                        Case 754, 872, 883
                            'Seguradoras sem nome - inválidas - não incluir.
                        Case Else
                            .AddItem Trim(rsP0044700!Seguradora)
                            .ItemData(.NewIndex) = val(rsP0044700!CSeguradora)
                        End Select
                        rsP0044700.MoveNext
                    Wend
                    If gM46V999.gTipoAcesso = DIGITACAO Then
                        .AddItem "**Não informado"
                        .ItemData(.NewIndex) = val(0)
                    End If
                End With
            End If
            Call gM46V999.gpFecha2(bdP0044700)
        End If
        panRenYasuda.Visible = False
        panRenMaritima.Visible = False
        panRenCongenere.Visible = True
    End Select

    Screen.MousePointer = vbDefault
    Exit Sub
ControleErro:
    Call gM46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gM46V999.gpGraLog(0, "T46V002A.optTipEmissao_Click ERRO - " & Err & " " & Err.Description)
End Sub

Private Sub txtRenApolice_GotFocus()
    With txtRenApolice
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub
Private Sub txtRenApolice_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub
Private Sub txtRenItem_GotFocus()
    With txtRenItem
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub
Private Sub txtRenItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

'1000347 - Inclusão de informações da congênere em renovações
Private Sub txtSegApolice_GotFocus()
    With txtSegApolice
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

'1000347 - Inclusão de informações da congênere em renovações
Private Sub txtSegApolice_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

'1000347 - Inclusão de informações da congênere em renovações
Private Sub txtSegItem_GotFocus()
    With txtSegItem
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

'1000347 - Inclusão de informações da congênere em renovações
Private Sub txtSegItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Carregar_CBO_SucursalMaritima()
    
    With cboSucursal
        .Clear
        .AddItem "Sucursal 01"
        .ItemData(.NewIndex) = 1
        .AddItem "Sucursal 02"
        .ItemData(.NewIndex) = 2
        .AddItem "Sucursal 03"
        .ItemData(.NewIndex) = 3
        .AddItem "Sucursal 04"
        .ItemData(.NewIndex) = 4
        .AddItem "Sucursal 05"
        .ItemData(.NewIndex) = 5
        .AddItem "Sucursal 08"
        .ItemData(.NewIndex) = 8
        .AddItem "Sucursal 09"
        .ItemData(.NewIndex) = 9
        .AddItem "Sucursal 14"
        .ItemData(.NewIndex) = 14
        .AddItem "Sucursal 18"
        .ItemData(.NewIndex) = 18
    End With
    
End Sub

Private Sub Carregar_CBO_RamoMaritima()
    
    With cboRamoMaritima
        .Clear
        .AddItem "Empresarial"
        .ItemData(.NewIndex) = 18
        .AddItem "Residencial"
        .ItemData(.NewIndex) = 14
        .AddItem "Condomínio"
        .ItemData(.NewIndex) = 16
    End With
    
End Sub

