VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmT47V021A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleciona tarifa/produto"
   ClientHeight    =   2625
   ClientLeft      =   150
   ClientTop       =   705
   ClientWidth     =   9000
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2625
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid flxProdutos 
      Height          =   1960
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   3466
      _Version        =   393216
      Rows            =   8
      Cols            =   8
      FixedCols       =   0
      ForeColor       =   12582912
      ForeColorFixed  =   12582912
      ForeColorSel    =   12582912
      WordWrap        =   -1  'True
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin Threed.SSCommand cmdOk 
      Default         =   -1  'True
      Height          =   300
      Left            =   3400
      TabIndex        =   0
      ToolTipText     =   "Volta à proposta."
      Top             =   2200
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Ok"
      ForeColor       =   8388608
   End
   Begin Threed.SSCommand cmdCancelar 
      Cancel          =   -1  'True
      Height          =   300
      Left            =   4600
      TabIndex        =   1
      ToolTipText     =   "Volta à proposta."
      Top             =   2200
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Cancelar"
      ForeColor       =   8388608
   End
End
Attribute VB_Name = "frmT47V021A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancelar_Click()
    Unload Me
End Sub
Private Sub cmdOK_Click()
    Screen.MousePointer = vbHourglass

    '1. Se cursor está posicionado na linha 0, nada foi selecionado.
    If flxProdutos.Row = 0 Then
        Screen.MousePointer = vbDefault
        Unload Me
        Exit Sub
    End If

    '2. Prepara controles.
    With flxProdutos
        .Col = 0
        gM46V999.gCodProduto = .Text
        .Col = 1
        gM46V999.gCodProtocolo = .Text
        .Col = 2
        gM46V999.gCodSequencia = .Text
    End With

    '3. Encerra processo.
    Screen.MousePointer = vbDefault
    Unload Me
End Sub
Private Sub flxProdutos_DblClick()
    Call cmdOK_Click
End Sub
Private Sub Form_Activate()
    If flxProdutos.Rows = 1 Then
        Unload Me
    End If
End Sub
Private Sub Form_Load()
    Dim bdTAB_PROD  As ADODB.Connection     'Arquivo - TAB_PROD.
    Dim lMensagem   As String               'Mensagem.
    Dim rsTAB_PROD  As clsYasRecordSet      'Registro - TAB_PROD.
Dim lSelect         As String
    Screen.MousePointer = vbHourglass

    '1. Posiciona form.
    Call gM46V999.gpCenForm(Me)

    '2. Prepara grade.
    With flxProdutos
        .Row = 0
        .Cols = 10
        'a) Posiciona largura das colunas.
        .ColWidth(0) = 0
        .ColWidth(1) = 0
        .ColWidth(2) = 0
        .ColWidth(3) = 2750
        .ColWidth(4) = 950
        .ColWidth(5) = 1550
        .ColWidth(6) = 950
        .ColWidth(7) = 2600
        .ColWidth(8) = 0
        .ColWidth(9) = 0
        'b) Posiciona alinhamento das colunas.
        .ColAlignment(4) = flexAlignCenterCenter
        .ColAlignment(5) = flexAlignCenterCenter
        .ColAlignment(6) = flexAlignCenterCenter
        .ColAlignment(7) = flexAlignCenterCenter
        'c) Posiciona títulos.
        .Col = 0
        .Text = "Código da tarifa"
        .Col = 1
        .Text = "Código do produto"
        .Col = 2
        .Text = "Número da sequência"
        .Col = 3
        .Text = "Nome do produto"
        .Col = 4
        .Text = "Tipo"
        .Col = 5
        .Text = "Corretor"
        .Col = 6
        .Text = "Versão"
        .Col = 7
        .Text = "Vigência"
        .Col = 8
        .Text = "Início da vigência - aaaammdd"
        .Col = 9
        .Text = "Término da vigência - aaaammdd"
    End With

    '3. Obtém registros de produtos e carrega grade.
    With flxProdutos
        If gM46V999.gfAbrBasDados("P00MULTI2", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, bdTAB_PROD, lMensagem) = False Then
            Call gM46V999.gpGraLog(2, lMensagem)
            Exit Sub
        End If
        .Rows = 1
        lSelect = "SELECT * FROM TAB_PRODUTO "
        lSelect = lSelect & " Where cod_ramo = " & gM46V999.gCod_Ramo
        If UCase(gM46V999.gSerDIVERSOS) Like "*MIX*" Then
            'não tem regra na pesquisa
        Else
            If gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Or gM46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda Then
                lSelect = lSelect & " and Tip_Prot = 1"
            End If
        End If
        lSelect = lSelect & " ORDER BY Cod_Produto DESC, Cod_Protocolo, Cod_Sequencia DESC"
        If gM46V999.gfObtRegistro(bdTAB_PROD, lSelect, rsTAB_PROD, lMensagem) = False Then
            Call gM46V999.gpFecha2(bdTAB_PROD)
            Call gM46V999.gpGraLog(2, lMensagem)
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        While Not rsTAB_PROD.EOF
            .Rows = .Rows + 1
            .Row = .Rows - 1
            .Col = 0
            .Text = rsTAB_PROD!Cod_Produto
            .Col = 1
            .Text = rsTAB_PROD!Cod_Protocolo
            .Col = 2
            .Text = rsTAB_PROD!Cod_Sequencia
            .Col = 3
            .Text = rsTAB_PROD!Nom_Prot
            .Col = 4
            .Text = rsTAB_PROD!Tip_Prot
            .Col = 5
            If IsNull(rsTAB_PROD!Cod_Corr_User) Then
                .Text = "---"
            Else
                .Text = rsTAB_PROD!Cod_Corr_User
            End If
            .Col = 6
            If IsNull(rsTAB_PROD!Cod_Versao_Prot) Then
                .Text = "---"
            Else
                .Text = rsTAB_PROD!Cod_Versao_Prot
            End If
            .Col = 7
            .Text = "de " & gM46V999.gfForData(2, rsTAB_PROD!Dat_Vig_Ini) & _
                    " a " & gM46V999.gfForData(2, rsTAB_PROD!Dat_Vig_Ter)
            .Col = 8
            .Text = rsTAB_PROD!Dat_Vig_Ini
            .Col = 9
            .Text = rsTAB_PROD!Dat_Vig_Ter
            rsTAB_PROD.MoveNext
        Wend
        Call gM46V999.gpFecha2(bdTAB_PROD)
        If .Rows = 1 Then       'Quantidade de registros.
            Call gM46V999.gpGraLog(2, "FTE2005 - " & gM46V999.gPrefixo & _
                             "Arquivo inválido (tabela de produtos - sem registro).")
        Else
            .Row = 1
        End If
        
    End With

    Screen.MousePointer = vbDefault
End Sub
