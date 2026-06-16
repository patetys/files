VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmT46V208A 
   Caption         =   "Inspeções Anteriores"
   ClientHeight    =   6435
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12780
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6435
   ScaleWidth      =   12780
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtVlRisco 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H80000015&
      Height          =   315
      Left            =   8850
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "500.000.000,00"
      Top             =   2100
      Width           =   3825
   End
   Begin VB.TextBox txtCnpj 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H80000015&
      Height          =   315
      Left            =   2370
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "00.000.000/0001-00"
      Top             =   1110
      Width           =   3825
   End
   Begin VB.ComboBox cboCondInsp 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "T46V208A.frx":0000
      Left            =   2370
      List            =   "T46V208A.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   1440
      Width           =   3855
   End
   Begin VB.TextBox txtInspObj 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H80000015&
      Height          =   315
      Left            =   8850
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "Sim"
      Top             =   1110
      Width           =   3825
   End
   Begin VB.TextBox txtCompl 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H80000015&
      Height          =   315
      Left            =   8850
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "Apto 1"
      Top             =   1770
      Width           =   3825
   End
   Begin VB.TextBox txtCep 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000015&
      Height          =   315
      Left            =   2370
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "04600-000"
      Top             =   2100
      Width           =   3825
   End
   Begin VB.TextBox txtItem 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H80000015&
      Height          =   315
      Left            =   2370
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "0001"
      Top             =   780
      Width           =   555
   End
   Begin VB.TextBox txtLocalRisco 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000015&
      Height          =   315
      Left            =   2370
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "R CUBATAO, 1"
      Top             =   1770
      Width           =   3825
   End
   Begin VB.TextBox txtNumInspPrg 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8850
      MaxLength       =   14
      TabIndex        =   0
      Text            =   "1234567"
      Top             =   1440
      Width           =   3825
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   12780
      _ExtentX        =   22543
      _ExtentY        =   1164
      ButtonWidth     =   1349
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgLista"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salvar"
            Key             =   "Salvar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Anterior"
            Key             =   "Anterior"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Próximo"
            Key             =   "Proximo"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair"
            Key             =   "Sair"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgLista 
      Left            =   8760
      Top             =   480
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
            Picture         =   "T46V208A.frx":0015
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V208A.frx":0367
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V208A.frx":06B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V208A.frx":0A0B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid flxInspecaoAnt 
      Height          =   3045
      Left            =   120
      TabIndex        =   15
      Top             =   2520
      Width           =   12555
      _ExtentX        =   22146
      _ExtentY        =   5371
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
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      Height          =   795
      Left            =   1920
      TabIndex        =   20
      Top             =   5520
      Width           =   9210
      Begin VB.Label Label14 
         BackColor       =   &H0000C000&
         Height          =   195
         Left            =   105
         TabIndex        =   36
         Top             =   480
         Width           =   195
      End
      Begin VB.Label Label13 
         BackColor       =   &H000000FF&
         Height          =   195
         Left            =   6840
         TabIndex        =   35
         Top             =   240
         Width           =   195
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FF80FF&
         Height          =   195
         Left            =   4530
         TabIndex        =   34
         Top             =   240
         Width           =   195
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FF0000&
         Height          =   195
         Left            =   2205
         TabIndex        =   33
         Top             =   240
         Width           =   195
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         Height          =   195
         Left            =   105
         TabIndex        =   32
         Top             =   240
         Width           =   200
      End
      Begin VB.Label LblLeg1 
         AutoSize        =   -1  'True
         Caption         =   "3 - Pendente Análise"
         Height          =   195
         Index           =   0
         Left            =   4830
         TabIndex        =   31
         ToolTipText     =   "Descrição do Status"
         Top             =   240
         Width           =   1470
      End
      Begin VB.Label LblLeg1 
         AutoSize        =   -1  'True
         Caption         =   "1 - Pendente Envio"
         ForeColor       =   &H80000006&
         Height          =   195
         Index           =   4
         Left            =   360
         TabIndex        =   30
         ToolTipText     =   "Descrição do Status"
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label LblLeg1 
         AutoSize        =   -1  'True
         Caption         =   "2 - Pendente Terceirizada"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   2520
         TabIndex        =   29
         ToolTipText     =   "Descrição do Status"
         Top             =   240
         Width           =   1830
      End
      Begin VB.Label LblLeg1 
         AutoSize        =   -1  'True
         Caption         =   "4 - Liberada Emissão"
         Height          =   195
         Index           =   7
         Left            =   7140
         TabIndex        =   28
         ToolTipText     =   "Descrição do Status"
         Top             =   240
         Width           =   1470
      End
      Begin VB.Label LblLeg1 
         AutoSize        =   -1  'True
         Caption         =   "5 - Inspeção Recusada"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   8
         Left            =   360
         TabIndex        =   27
         ToolTipText     =   "Descrição do Status"
         Top             =   480
         Width           =   1665
      End
      Begin VB.Label LblLeg1 
         Caption         =   "8 - Inspeção Liberada"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   11
         Left            =   7140
         TabIndex        =   26
         ToolTipText     =   "Descrição do Status"
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackColor       =   &H00404040&
         Height          =   195
         Index           =   0
         Left            =   6840
         TabIndex        =   25
         Top             =   480
         Width           =   195
      End
      Begin VB.Label Label7 
         BackColor       =   &H0000C0C0&
         Height          =   195
         Index           =   1
         Left            =   2205
         TabIndex        =   24
         Top             =   480
         Width           =   195
      End
      Begin VB.Label LblLeg1 
         AutoSize        =   -1  'True
         Caption         =   "6 - Inspeção Emitida"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   13
         Left            =   2520
         TabIndex        =   23
         ToolTipText     =   "Descrição do Status"
         Top             =   480
         Width           =   1440
      End
      Begin VB.Label Label7 
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   4530
         TabIndex        =   22
         Top             =   480
         Width           =   195
      End
      Begin VB.Label LblLeg1 
         AutoSize        =   -1  'True
         Caption         =   "7 - Inspeção Cancelada"
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   14
         Left            =   4830
         TabIndex        =   21
         ToolTipText     =   "Descrição do Status"
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Valor de Risco : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6600
      TabIndex        =   19
      Top             =   2100
      Width           =   2235
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CPF/CNPJ : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   120
      TabIndex        =   17
      Top             =   1110
      Width           =   2235
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Inspeção Obrigatória : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6600
      TabIndex        =   10
      Top             =   1110
      Width           =   2235
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cep : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   120
      TabIndex        =   9
      Top             =   2100
      Width           =   2235
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Complemento : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6600
      TabIndex        =   8
      Top             =   1770
      Width           =   2235
   End
   Begin VB.Label lblItem 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Item : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   780
      Width           =   2235
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Condições de Inspeção : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   2235
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Local de Risco : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   1770
      Width           =   2235
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Inspeção Programada : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6600
      TabIndex        =   4
      Top             =   1440
      Width           =   2235
   End
End
Attribute VB_Name = "frmT46V208A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private colPedloc As Collection
Private colInspecao As Collection
Private intPosicao As Integer
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private pStatus As Byte
Private PCodCor As String

Private Sub ConfigGridCbo()
    With flxInspecaoAnt
        .Cols = 11
        .Rows = 1
                       
        .Col = 0
        .TextMatrix(0, 0) = "Item"
        .ColWidth(0) = 400
        .TextMatrix(0, 1) = "Inspeção"
        .ColWidth(1) = 800
        .TextMatrix(0, 2) = "Local de Risco"
        .ColWidth(2) = 2500
        .TextMatrix(0, 3) = "Complemento"
        .ColWidth(3) = 1300
        .TextMatrix(0, 4) = "CEP"
        .ColWidth(4) = 1000
        .TextMatrix(0, 5) = "Bairro"
        .ColWidth(5) = 1700
        .TextMatrix(0, 6) = "Cidade"
        .ColWidth(6) = 1300
        .TextMatrix(0, 7) = "UF"
        .ColWidth(7) = 300
        .TextMatrix(0, 8) = "Data Parecer"
        .ColWidth(8) = 1300
        .TextMatrix(0, 9) = "Usuário"
        .ColWidth(9) = 800
        .TextMatrix(0, 10) = "Classificação"
        .ColWidth(10) = 1100
        .Col = 0
    End With

    txtNumInspPrg.Enabled = False
    cboCondInsp.Clear
    cboCondInsp.AddItem "Conforme Normas"
    cboCondInsp.ItemData(cboCondInsp.NewIndex) = e_StatusInspecao_ConfNormas
    cboCondInsp.AddItem "Dispensada"
    cboCondInsp.ItemData(cboCondInsp.NewIndex) = e_StatusInspecao_Dispensada
    cboCondInsp.AddItem "Obrigatória"
    cboCondInsp.ItemData(cboCondInsp.NewIndex) = e_StatusInspecao_obrigatoria
    cboCondInsp.AddItem "Programada"
    cboCondInsp.ItemData(cboCondInsp.NewIndex) = e_StatusInspecao_programada
    
End Sub

Public Sub CorDetalhe()

    Select Case pStatus
        Case 1
            PCodCor = &HC0C000
        Case 2
            PCodCor = &HFF0000
        Case 3
            PCodCor = &HFF80FF
        Case 4
            PCodCor = &HFF
        Case 5
            PCodCor = &HC000&
        Case 6
            PCodCor = &HC0C0&
        Case 7
            PCodCor = &H80C0FF
        Case 8
            PCodCor = &H404040
    End Select

End Sub

Public Sub CarregarForm(ByRef pcolPedloc As Collection, pcolInspecao As Collection, ByVal pPosicao As Integer)
    Call ConfigGridCbo
    Set colPedloc = pcolPedloc
    Set colInspecao = pcolInspecao
    intPosicao = pPosicao
    Call Posiciona
    Me.Show vbModal
End Sub

Private Function Posiciona() As Boolean
Dim lI As Integer
Dim stcobjInspecao As stcInspecao
    On Error GoTo Erro

    Posiciona = False
    With colPedloc(intPosicao)
        txtItem.Text = .Num_Item
        txtCnpj.Text = .ObjPai.gCNPJ_CPF_Formatado
        For lI = 0 To cboCondInsp.ListCount - 1
            If cboCondInsp.ItemData(lI) = .Ind_StatusInspecao Then
                cboCondInsp.ListIndex = lI
                Exit For
            End If
        Next

        txtNumInspPrg.Text = .Num_Insp
        txtLocalRisco.Text = .Cod_Tip_Logr_Risco & " " & .Nom_Logr_Risco & ", " & .objstcPedInsp.Num_Logr_Risco
        txtCompl.Text = .Dsc_Compl_Risco
        txtCep.Text = .gCEP_Formatado '.Num_Cep_Risco
        txtVlRisco.Text = Format(.Val_Vr, "###,###,##0.00")
        
        txtInspObj.Text = IIf(.Insp_Obr = True, "Sim", "Não")
    End With
    
    If Not (colInspecao Is Nothing) Then
        flxInspecaoAnt.Rows = 1
        lI = 1
        Set stcobjInspecao = New stcInspecao
        For Each stcobjInspecao In colInspecao
            With stcobjInspecao
                If .Num_Ind = intPosicao Then
                    pStatus = .cod_status
                    CorDetalhe
                    flxInspecaoAnt.Rows = flxInspecaoAnt.Rows + 1
                    flxInspecaoAnt.Row = lI
                    flxInspecaoAnt.Col = 0
                    flxInspecaoAnt.CellForeColor = PCodCor
                    flxInspecaoAnt.TextMatrix(lI, 0) = .Num_Item       '"Item"
                    flxInspecaoAnt.Col = 1
                    flxInspecaoAnt.CellForeColor = PCodCor
                    flxInspecaoAnt.TextMatrix(lI, 1) = .Num_Insp       '"Inspeção"
                    flxInspecaoAnt.Col = 2
                    flxInspecaoAnt.CellForeColor = PCodCor
                    flxInspecaoAnt.TextMatrix(lI, 2) = RTrim(.tip_logr) & " " & RTrim(.Nom_Logr) & ", " & .Num_Logr       '"Local de Risco"
                    flxInspecaoAnt.Col = 3
                    flxInspecaoAnt.CellForeColor = PCodCor
                    flxInspecaoAnt.TextMatrix(lI, 3) = .dsc_compl      '"Complemento"
                    flxInspecaoAnt.Col = 4
                    flxInspecaoAnt.CellForeColor = PCodCor
                    
                    flxInspecaoAnt.TextMatrix(lI, 4) = Format(Format(.Num_CEP, "00000000"), "00000-000")      '"CEP"
                    flxInspecaoAnt.Col = 5
                    flxInspecaoAnt.CellForeColor = PCodCor
                    flxInspecaoAnt.TextMatrix(lI, 5) = .dsc_bairro     '"Bairro"
                    flxInspecaoAnt.Col = 6
                    flxInspecaoAnt.CellForeColor = PCodCor
                    flxInspecaoAnt.TextMatrix(lI, 6) = .dsc_cid        '"Cidade"
                    flxInspecaoAnt.Col = 7
                    flxInspecaoAnt.CellForeColor = PCodCor
                    flxInspecaoAnt.TextMatrix(lI, 7) = .uf             '"UF"
                    flxInspecaoAnt.Col = 8
                    flxInspecaoAnt.CellForeColor = PCodCor
                    flxInspecaoAnt.TextMatrix(lI, 8) = gM46V999.gfForData(2, .dat_derie)       '"Data Parecer"
                    flxInspecaoAnt.Col = 9
                    flxInspecaoAnt.CellForeColor = PCodCor
                    flxInspecaoAnt.TextMatrix(lI, 9) = .cod_user_cad    '"Usuário"
                    flxInspecaoAnt.Col = 10
                    flxInspecaoAnt.CellForeColor = PCodCor
                    flxInspecaoAnt.TextMatrix(lI, 10) = .Cod_Clasf & "-" & .Cod_Compl_Clasf    '"Classificação"
                    flxInspecaoAnt.Col = 0
                    lI = lI + 1
                End If
            End With
        Next stcobjInspecao
    End If
    
    Posiciona = True
    Exit Function
Erro:
    MsgBox Err.Number & " - " & Err.Description
End Function

Private Sub Anterior()

    If Not Salvar Then Exit Sub
    
    intPosicao = intPosicao - 1
    If intPosicao < 1 Then
        intPosicao = 1
    Else
        Call Posiciona
    End If
    
End Sub

Private Sub Proximo()

    If Not Salvar Then Exit Sub
    
    intPosicao = intPosicao + 1
    If intPosicao > colPedloc.Count Then
        intPosicao = colPedloc.Count
    Else
        Call Posiciona
    End If
End Sub

Private Function Salvar() As Boolean
Dim lobjPedLoc As stcA46V708B
Dim objConexao As ADODB.Connection
Dim pMensagem As String
    On Error GoTo Erro

    Salvar = False
    
    If cboCondInsp.ListIndex = e_StatusInspecao_programada Then
        If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, objConexao, pMensagem) = False Then
            Call gM46V999.gpGraLog(2, pMensagem)
            Exit Function
        End If
        With colPedloc(intPosicao)
            If Not gM46V999.gfValidaInspecao(.Num_Item, txtNumInspPrg.Text, .ObjPai.Num_Cgc_Cpf, .Num_Cep_Risco, .objstcPedInsp.Num_Logr_Risco, pMensagem, objConexao) Then
                Call gM46V999.gpGraLog(1, pMensagem)
                cboCondInsp.SetFocus
                Exit Function
            End If
        End With
    End If

    Set lobjPedLoc = New stcA46V708B
    Set lobjPedLoc = colPedloc(intPosicao)
            
    With colPedloc(intPosicao)
        .Ind_StatusInspecao = cboCondInsp.ListIndex
        .Num_Insp = txtNumInspPrg.Text
    End With
    
    Salvar = True
    Exit Function
Erro:
    MsgBox Err.Number & " - " & Err.Description
End Function

Private Sub Form_Load()
    Call gM46V999.gpCenForm(Me)
End Sub

Private Sub cboCondInsp_Click()
    txtNumInspPrg.Enabled = False
    txtNumInspPrg.Text = 0
    If cboCondInsp.ListIndex = e_StatusInspecao_programada Then
        txtNumInspPrg.Enabled = True
    End If
End Sub

Private Sub flxInspecaoAnt_Click()
    Clipboard.Clear
    Clipboard.SetText (flxInspecaoAnt.TextMatrix(flxInspecaoAnt.Row, 1))
End Sub

Private Sub flxInspecaoAnt_DblClick()
    Call VisualizaInsp
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = 0 Then
        Cancel = 1
        MsgBox "Utilize o botão sair!", vbInformation, Me.Caption
        Exit Sub
    End If
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Salvar"
            If Salvar Then
                Unload Me
            End If
        Case "Anterior"
            Call Anterior
        Case "Proximo"
            Call Proximo
        Case "Sair"
            Unload Me
    End Select
End Sub

Private Sub txtNumInspPrg_KeyPress(KeyAscii As Integer)
    KeyAscii = ConsisteKeyNum(KeyAscii, ",")

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub VisualizaInsp()
Dim lArqInspec          As String
Dim lDirArq             As String
Dim lNumSeq             As Integer
Dim NumInsp             As String
Dim NomedoArquivo       As String
Dim Sql                 As String
Dim lMensagem           As String
Dim lrsLaudo            As ADODB.Recordset
Dim lbdTab_Insp         As ADODB.Connection
On Error GoTo Err_Acesso

    If flxInspecaoAnt.TextMatrix(flxInspecaoAnt.Row, 1) = "Inspeção" Then
       Exit Sub
    End If
 
    NumInsp = (flxInspecaoAnt.TextMatrix(flxInspecaoAnt.Row, 1))
    Ampulheta True

     If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTab_Insp, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If
    
    Set lrsLaudo = New ADODB.Recordset
    Sql = "Select dat_agend from "
    Sql = Sql & " Inspecao.dbo.Tab_Insp (nolock)"
    Sql = Sql & " where  Num_insp  = " & NumInsp
    lrsLaudo.Open Sql, lbdTab_Insp
    If Not lrsLaudo.EOF Then
        If lrsLaudo("dat_agend") <> 0 Then
            NomedoArquivo = "http://www.yasuda.com.br/YasForms/Controller?action=consultar&Cod_Iden_Doc=" & NumInsp
            ShellExecute Me.hWnd, "Open", NomedoArquivo, vbNullString, "C:\", 1
            Ampulheta False
        Else
            MsgBox "Inspeção nr " & NumInsp & " não disponível ainda ", vbExclamation, "Visualização da Inspeção"
        End If
    End If
    Call gM46V999.gpFecha2(lbdTab_Insp)
    Ampulheta False

Exit Sub
Err_Acesso:
  MsgBox "Erro na visualização do Laudo, acionar o sistemas " & Err & " - " & Error, , "Inspeção"
  Exit Sub
  Resume Next
End Sub

Sub Ampulheta(Modo As Integer)
    If Modo = -1 Then
        Screen.MousePointer = 11
    ElseIf Modo = 3 Then
    Else
        Screen.MousePointer = 0
    End If
End Sub
