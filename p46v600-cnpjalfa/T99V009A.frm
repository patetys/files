VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmT99V009A 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Geração do boleto - antecipado."
   ClientHeight    =   6510
   ClientLeft      =   1860
   ClientTop       =   2190
   ClientWidth     =   10470
   Icon            =   "T99V009A.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   10470
   Begin VB.VScrollBar VScroll1 
      Height          =   5895
      LargeChange     =   500
      Left            =   10200
      Max             =   8600
      TabIndex        =   3
      Top             =   480
      Width           =   255
   End
   Begin Threed.SSPanel panBoleto 
      Height          =   5895
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   10455
      _Version        =   65536
      _ExtentX        =   18441
      _ExtentY        =   10398
      _StockProps     =   15
      BackColor       =   13160660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      Begin SHDocVwCtl.WebBrowser wbBoleto 
         Height          =   15975
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   10455
         ExtentX         =   18441
         ExtentY         =   28178
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
   End
   Begin MSComctlLib.ImageList lstImagens 
      Left            =   3480
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T99V009A.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T99V009A.frx":0326
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T99V009A.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T99V009A.frx":0A92
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T99V009A.frx":0EE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T99V009A.frx":11FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T99V009A.frx":1518
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T99V009A.frx":1832
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T99V009A.frx":1B4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T99V009A.frx":1E66
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T99V009A.frx":2180
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T99V009A.frx":249A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T99V009A.frx":27B4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar barFerramentas 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "lstImagens"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Imprimir"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Menu mnuArquivo 
      Caption         =   "Arquivo"
      Begin VB.Menu mnuArqImprimir 
         Caption         =   "Im&primir"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArqSair 
         Caption         =   "&Sair"
      End
   End
End
Attribute VB_Name = "frmT99V009A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mEndArqNovo As String
Dim lIniciouFluxo As Boolean

Public Sub gpInstanciarForm(pobjstc As stcA99V009B, pShow As Boolean)
    Dim lEndTemplate As String
    Dim eQuery As OLECMDF
    Dim objBoleto As clsA99V009A

    Set objBoleto = New clsA99V009A
    If Not objBoleto.mfMontarBoleto(pobjstc) Then
        Call gM46V999.gpGraLog(1, objBoleto.UltimoErro)
        Exit Sub
    End If
    lEndTemplate = App.Path & "\P0045200.exe"
    mEndArqNovo = App.Path & "\" & pobjstc.NossoNumero & ".HTML"
    If Not objBoleto.mfRetornarPaginaMontada(pobjstc, lEndTemplate, mEndArqNovo) Then
        Call gM46V999.gpGraLog(1, objBoleto.UltimoErro)
        Exit Sub
    End If


    If objBoleto.UltimoErro <> "" Then
        Call gM46V999.gpGraLog(1, objBoleto.UltimoErro)
    Else
        wbBoleto.Navigate2 mEndArqNovo
    End If
    On Error Resume Next
    Kill mEndArqNovo
    On Error GoTo 0

    If pShow Then
        Screen.MousePointer = vbDefault
        Me.Show vbModal
    Else
        Call mpImprimir
    End If
    On Error Resume Next
    Kill mEndArqNovo
    On Error GoTo 0
End Sub
Private Sub teste()
    Dim objstc As stcA99V009B
    Dim obj As clsA99V009A
    Dim lConexao As ADODB.Connection

    Set obj = New clsA99V009A
    Set objstc = New stcA99V009B
    Set objstc.M46V999 = gM46V999

    With objstc
        Set lConexao = New ADODB.Connection
        lConexao.ConnectionString = "Provider=MSDASQL.1;Extended Properties=""DRIVER=SQL Server;SERVER=SPX20112CSLUMIX;UID=workflow;PWD=yasworkflow;APP=Visual Basic;WSID=WPD10600FSLU166;DATABASE=Transporte"""
        lConexao.Open
        Set .ConexaoBD = lConexao
        .NomeTabela = "yasudanet.dbo.Tab_Contas_Carne"
        .CarregarDados = True
        .vencimento = 20070916
        .NossoNumero = 5703223
        .NossoNumeroMask = "5703223-4"
        .valorDocumento = 241.16
        .numeroBanco = 341

        .NomeSacado = "IRAIDE"
        .Vigencia = "30/08/2007 a 20/08/2008"
        .LocalImagensBarra = App.Path & "/boleto_arquivos"
        .LocalImagensBanco = App.Path & "/boleto_arquivos/logo_bradesco_boleto.gif"
        .Corretor = "NOVAES"
        .MensagemPgto = "251 - 1601 PARC 01/01"
        .ramo = "310 - Automóvel"
    End With

    Call obj.mfMontarBoleto(objstc)
    Call obj.mfRetornarPaginaMontada(objstc, "C:\User_c\Projetos\P99V009\testecodigo.exe", "C:\User_c\Projetos\P99V009\testecodigo.html")


    If obj.UltimoErro <> "" Then
        MsgBox obj.UltimoErro
    Else
        wbBoleto.Navigate2 "C:\User_c\Projetos\P99V009\testecodigo.html"
    End If

End Sub
Private Sub mpImprimir()
    Dim eQuery As OLECMDF
    eQuery = wbBoleto.QueryStatusWB(OLECMDID_PRINT)
    If eQuery And OLECMDF_ENABLED Then
        wbBoleto.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER
        On Error Resume Next
        Kill mEndArqNovo
        On Error GoTo 0
    End If
End Sub
Private Sub barFerramentas_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case UCase(Button.Key)
    Case "IMPRIMIR"
        Call mpImprimir
        Unload Me
    Case "SAIR"
        Unload Me
    End Select
End Sub

Private Sub Form_Load()
    Call gM46V999.gpCenForm(Me)
End Sub

Private Sub mnuArqImprimir_Click()
    Call mpImprimir
End Sub

Private Sub mnuArqSair_Click()
    Unload Me
End Sub

Private Sub VScroll1_Change()
    Call VScroll1_Scroll
End Sub

Private Sub VScroll1_Scroll()
    wbBoleto.Top = VScroll1.Value * -1
End Sub

Private Sub wbBoleto_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    VScroll1.Value = 1
    Screen.MousePointer = vbDefault
    Me.MousePointer = vbDefault
End Sub


