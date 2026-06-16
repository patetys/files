VERSION 5.00
Begin VB.Form frmT46V198A 
   Caption         =   "Propostas Transmitidas (originais)"
   ClientHeight    =   3750
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5340
   LinkTopic       =   "Pesquisa Propostas Transmitidas"
   ScaleHeight     =   3750
   ScaleWidth      =   5340
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdAbrir 
      Caption         =   "Abrir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   320
      TabIndex        =   6
      Top             =   2640
      Width           =   1695
   End
   Begin VB.OptionButton OptionPesquisa 
      Caption         =   "Número PI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   0
      Left            =   320
      TabIndex        =   1
      Top             =   360
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.TextBox txtNossoNumero 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      MaxLength       =   20
      TabIndex        =   5
      Top             =   840
      Width           =   2400
   End
   Begin VB.OptionButton OptionPesquisa 
      Caption         =   "Nosso Número"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   1
      Left            =   320
      TabIndex        =   4
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox txtDigito 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      MaxLength       =   1
      TabIndex        =   3
      Top             =   360
      Width           =   350
   End
   Begin VB.CommandButton CmdLimpar 
      Caption         =   "Limpar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3350
      TabIndex        =   8
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox txtNumeroPI 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      MaxLength       =   10
      TabIndex        =   2
      Top             =   360
      Width           =   1750
   End
   Begin VB.Frame FramePesquisa 
      Caption         =   "   Pesquisar     "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   3735
      Left            =   80
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin VB.CommandButton CmdImprimir 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Gerar   Impressão"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   3120
         Visible         =   0   'False
         Width           =   4720
      End
      Begin VB.Label lblMensagem 
         AutoSize        =   -1  'True
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   120
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   4080
         TabIndex        =   10
         Top             =   360
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmT46V198A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''Private gM46V999 As clsM46V999
Private clsA99V001 As clsA99V001

Dim bdBasDados As ADODB.Connection
Dim lObjPed As clsA46V702A
Dim lObjstcPed As stcA46V702B
Dim lMensagem As String
Dim lDigito As String
Dim bHabImpressao As Boolean

Private gOrigem As Integer

Private Sub Form_Load()
    IniciarControles
End Sub

Private Sub cmdImprimir_Click()
    Dim RelatoriosWebService As MSSOAPLib30.SoapClient30
    Dim EnderecoWebService As String
    Dim lUsuario As String
    Dim lLocal As String
    Dim lArquivo() As Byte
    Dim fnum As Integer
    
    Screen.MousePointer = vbArrowHourglass
    
    On Error GoTo Erro

    '' Configura WebService
    EnderecoWebService = gM46V999.gUrlRelatorio
    If EnderecoWebService = Empty Then
        If InStr(UCase(Command), "DEV") = 0 Then
            EnderecoWebService = "http://192.168.1.93:9080/SyasWeb02/WSMulti/SyasMultiRelatorioServices.svc?singleWsdl"     'PROD
        Else
            EnderecoWebService = "http://192.168.1.92:9080/SyasWeb01/WSMulti/SyasMultiRelatorioServices.svc?singleWsdl"     'DEV
        End If
    End If
    
    lUsuario = gM46V999.gUsuarioRelatorio
    If lUsuario = Empty Then
        If InStr(UCase(Command), "DEV") = 0 Then
            lUsuario = "ysnt09"     'PROD
        Else
            lUsuario = "test01"     'DEV
        End If
    End If
    
    Set RelatoriosWebService = New MSSOAPLib30.SoapClient30
    RelatoriosWebService.MSSoapInit (EnderecoWebService)
    RelatoriosWebService.ClientProperty("ServerHTTPRequest") = True
    RelatoriosWebService.ConnectorProperty("Timeout") = 900000        '15 min (com 200 itens leva 95seg em dev)
    RelatoriosWebService.ConnectorProperty("ConnectTimeout") = 900000 '15 min

    lArquivo = RelatoriosWebService.RetornarRelatorioFTP(lObjstcPed.NOSSO_NUMERO, lUsuario, gM46V999.gNomUsuario, gOrigem)
   
    'Salvar o Arquivo PDF
    lLocal = "C:\Temp\"
    lLocal = lLocal & "MULT - " & lObjstcPed.Num_Ped & ".PDF"
        
    Dim i
    i = 0
    Do While Len(Dir(lLocal)) > 0
        i = i + 1
        If InStr(lLocal, "(") > 0 Then
            lLocal = Left(lLocal, InStr(lLocal, "(") - 1) & "(" & i & ").PDF"
        Else
            lLocal = Left(lLocal, InStr(lLocal, ".") - 1) & "(" & i & ").PDF"
        End If
    Loop
       
    fnum = FreeFile
    Open lLocal For Binary As #fnum
    Put #fnum, 1, lArquivo
    Close fnum

    frmPDF.CaminhoRelatorio = lLocal
    frmPDF.Show vbModal
    
    'Mensagem informando local do arquivo salvo
    If Trim(Dir(lLocal)) <> Empty And lLocal <> Empty Then
        Call gM46V999.gpGraLog(1, "Arquivo: " & lLocal & " gerado com sucesso!")
    End If

Fim:
    Screen.MousePointer = vbDefault
    Exit Sub

Erro:
    MsgBox "WebService erro: " & Err.Description, vbCritical, Me.Caption
    Resume Fim
    Resume 0
End Sub

Private Sub OptionPesquisa_Click(index As Integer)
    CheckOption
End Sub

Private Sub CheckOption()
    If OptionPesquisa(0).Value = True Then
        txtNumeroPI.Enabled = True
        txtDigito.Enabled = True
        txtNossoNumero.Enabled = False
        txtNossoNumero.Text = ""
    ElseIf OptionPesquisa(1).Value = True Then
        txtNossoNumero.Enabled = True
        txtNumeroPI.Enabled = False
        txtDigito.Enabled = False
        txtNumeroPI.Text = ""
        txtDigito.Text = ""
    End If
End Sub

Private Sub txtNumeroPI_Change()
    CalcDigito
    lblMensagem.Caption = ""
End Sub

Private Sub CalcDigito()
    Set clsA99V001 = New clsA99V001
    Set clsA99V001.M46V999 = gM46V999

    If txtNumeroPI.Text <> "" Then
        If clsA99V001.DIG_YASUDA(txtNumeroPI.Text, lDigito, lMensagem) = False Then
            Exit Sub
        End If
        
        txtDigito.Text = lDigito
    End If
    Set clsA99V001 = Nothing
End Sub

Private Sub CmdAbrir_Click()
    '' Limpar variaveis/mensagem
    lblMensagem.Caption = ""
    bHabImpressao = False
    Screen.MousePointer = vbHourglass
    
    Dim vCodUnicidade As Long
    vCodUnicidade = 0
    lMensagem = ""
    
    'On Error GoTo Sair
    
    '' Validacoes basicas
    If (OptionPesquisa(0).Value = True) Then
        If Len(Trim(txtNumeroPI.Text)) <= 7 Then
            txtNumeroPI.Text = Right(Year(Date), 2) & "2" & Format(txtNumeroPI.Text, "0000000")
        Else
            txtNumeroPI.Text = Format(txtNumeroPI.Text, "0000000000")
        End If
        
        If IsNumeric(txtNumeroPI) Then
            CalcDigito
        Else
            lblMensagem.Caption = "Número PI inválido!"
            GoTo Sair
        End If
    End If

    If (OptionPesquisa(1).Value = True) Then
        If Not IsNumeric(txtNossoNumero) Or Len(txtNossoNumero.Text) <> 20 Then
            lblMensagem.Caption = "Nosso Número inválido!"
            GoTo Sair
        End If
    End If
     
    '' Testar conexao base
    If gM46V999.gfAbrBasDados("P00MPED", gM46V999.gTipP00MPED, gM46V999.gEndP00MPED, bdBasDados, lMensagem) = False Then
        Call gM46V999.gpGraLog(1, "FNP0044 - " & lMensagem)
        lblMensagem.Caption = lMensagem     '' Mensagem Erro
        GoTo Sair
    End If
    
    '' Buscar
    Set lObjPed = New clsA46V702A
    Set lObjPed.M46V999 = gM46V999
    Set lObjPed.ConexaoBD = bdBasDados
    
    '' Buscar Unicidade
    Call gM46V999.gfBuscaNossoNumeroUnicidade("FTP_", txtNumeroPI.Text, txtNossoNumero.Text, vCodUnicidade, lMensagem)
   
    Set lObjstcPed = New stcA46V702B
    
    If OptionPesquisa(0).Value = True Then
        lObjstcPed.NUM_PI = txtNumeroPI.Text
    Else
        lObjstcPed.NOSSO_NUMERO = txtNossoNumero.Text
    End If
    
    lObjstcPed.Cod_Unicidade = vCodUnicidade


    '' Carregar da FTP
    gOrigem = 3
    If Not lObjPed.mfCarregarColecao(lObjstcPed, "", "FTP_") Then
        Call gM46V999.gpGraLog(1, "FNP0043 - (FTP) " & lObjPed.UltimoErro)
        GoTo Sair
    End If

    '' 2016-08-31: Comentado para nao trazer
    '' Carregar do Ramos
    'If lObjPed.Colecao.Count = 0 Then
    '    gOrigem = 1
        
    '    If Not lObjPed.mfCarregarColecao(lObjstcPed, "RamosDiversos..", "") Then
    '        Call gM46V999.gpGraLog(1, "FNP0043 - (Ramos) " & lObjPed.UltimoErro)
    '        GoTo Sair
    '    End If
    'End If

    '' Popula...
    If lObjPed.Colecao.Count <> 0 Then
        '' Habilitar botao impresao
        bHabImpressao = True
        
        '' Popular texto
        Set lObjstcPed = lObjPed.Colecao.item(1)
        
        If OptionPesquisa(0).Value = True Then
            txtNossoNumero.Text = lObjstcPed.NOSSO_NUMERO
        Else
            txtNumeroPI.Text = lObjstcPed.NUM_PI
        End If
        
        lblMensagem.Caption = "Ramo......:" + Space(2) + CStr(lObjstcPed.Cod_Ramo) + vbCrLf _
                            + "Referência:" + Space(2) + CStr(lObjstcPed.Num_Ped) + vbCrLf _
                            + "Segurado..:" + Space(2) + lObjstcPed.Nom_Segurado + vbCrLf _
                            + "Digitação.:" + Space(2) + FormataData(lObjstcPed.DAT_DIGIT) + vbCrLf _
                            + "Vigência..:" + Space(2) + FormataData(lObjstcPed.Dat_Ini_Vig) + " ~ " + FormataData(lObjstcPed.Dat_Ter_Vig)
    Else
        lblMensagem.Caption = "Proposta não encontrada !"
    End If
    
Sair:
    Call gM46V999.gpFecha2(bdBasDados)
    Screen.MousePointer = vbDefault
    CmdImprimir.Visible = bHabImpressao
End Sub

Private Function FormataData(val As Long) As String
    Dim strAux As String
    
    If val < 19000101 Then
        val = 19000101
    End If

    strAux = CStr(val)
    
    FormataData = Mid$(strAux, 7, 2) & "/" & Mid$(strAux, 5, 2) & "/" & Mid$(strAux, 1, 4)
End Function

Private Sub CmdLimpar_Click()
    IniciarControles
End Sub

Private Sub IniciarControles()
    gOrigem = 3
    OptionPesquisa(0).Value = True
    
    txtNumeroPI.Enabled = True
    txtDigito.Enabled = True
    txtNumeroPI.Text = ""
    txtDigito.Text = ""
        
    txtNossoNumero.Enabled = False
    txtNossoNumero.Text = ""
    
    lblMensagem.Caption = ""
    
    bHabImpressao = False
    CmdImprimir.Visible = False
End Sub

