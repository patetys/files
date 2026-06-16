VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmT46V211A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Observação"
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7755
   ControlBox      =   0   'False
   Icon            =   "T46V211A.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   7755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Fechar"
      Height          =   400
      Left            =   4920
      TabIndex        =   4
      Top             =   7680
      Width           =   1400
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "Salvar"
      Height          =   400
      Left            =   1320
      TabIndex        =   3
      Top             =   7680
      Width           =   1400
   End
   Begin VB.Frame frmArquivos 
      Caption         =   "   Documentos anexados   "
      Height          =   3255
      Left            =   120
      TabIndex        =   1
      Top             =   4200
      Width           =   7455
      Begin MSFlexGridLib.MSFlexGrid flxArquivos 
         Height          =   2175
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   3836
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColor       =   -2147483644
         ForeColor       =   8388608
         BackColorFixed  =   -2147483636
         ForeColorFixed  =   16777215
         BackColorSel    =   65535
         ForeColorSel    =   12648447
         BackColorBkg    =   -2147483633
         GridColorFixed  =   4210752
         AllowBigSelection=   0   'False
         GridLines       =   0
         GridLinesFixed  =   0
         ScrollBars      =   2
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
      Begin VB.PictureBox pictArqAnexo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   5160
         Picture         =   "T46V211A.frx":000C
         ScaleHeight     =   342.857
         ScaleMode       =   0  'User
         ScaleWidth      =   300
         TabIndex        =   7
         Top             =   360
         Visible         =   0   'False
         Width           =   300
      End
   End
   Begin VB.Frame frmTexto 
      Caption         =   "   Texto digitado pelo Corretor ou Subscritor   "
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      Begin VB.TextBox txtObservacao 
         BackColor       =   &H00FFFFFF&
         CausesValidation=   0   'False
         Height          =   3495
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   6975
      End
   End
   Begin VB.Label lblMensagem 
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   8280
      Width           =   4935
   End
End
Attribute VB_Name = "frmT46V211A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const SW_HIDE As Long = 0
Private Const SW_SHOWNORMAL As Long = 1
Private Const SW_SHOWMAXIMIZED As Long = 3
Private Const SW_SHOWMINIMIZED As Long = 2

Public Enum T46V211A_Acao
    Consultar = 1
    Editar = 2
    Incluir = 3
End Enum

Public Acao As T46V211A_Acao
Private objBCol As New Collection
    
Private mObj750 As stcA46V750B
Private mObj753 As stcA46V753B

Private Sub Class_Initialize()
    Set mObj750 = New stcA46V750B
    Set mObj753 = New stcA46V753B
End Sub

Private Sub Class_Terminate()
    Set mObj750 = Nothing
    Set mObj753 = Nothing
End Sub

Public Property Get ObjStc750B() As stcA46V750B
    Set ObjStc750B = mObj750
End Property

Public Property Set ObjStc750B(obj As stcA46V750B)
    Set mObj750 = obj
End Property

Public Property Get ObjStc753B() As stcA46V753B
    Set ObjStc753B = mObj753
End Property

Public Property Set ObjStc753B(obj As stcA46V753B)
    Set mObj753 = obj
End Property

                    
Private Sub Form_Load()
    'If ObjStc750B Is Nothing Then
    '   '' Erro nos parametros, sair
    '   Call Sair
    'End If
     
    '' Carregar os dados.
''    Dim obj753 As New stcA46V753B, obj750 As New stcA46V750B
''    Set obj753 = ObjStc753B
     
    txtObservacao.Text = mObj750.Observacao
    
    If Acao = Consultar Then
        cmdSalvar.Enabled = False
        txtObservacao.Enabled = False
    End If
    
    Call CarregarAnexos(mObj750)
    
    Call gM46V999.gpCenForm(Me)
End Sub

Private Sub CarregarAnexos(objFlex As stcA46V750B)
    Dim i As Integer, msg As String
    Dim objA As New clsA46V752A
    Dim objB As New stcA46V752B
    
    '' Conexao e variaveis globais
    Set objA.M46V999 = gM46V999
    
    objB.IdObs = objFlex.IdObs
    
    '' Sem anexos: fim
    If objA.mfCarregarColecao(objB) = False Then
        GoTo Fim
    End If

    '' Colecao de anexos
    Set objBCol = objA.Colecao

    If objBCol.Count <= 0 Then
        lblMensagem.Caption = "Sem anexos disponíveis."
        flxArquivos.Visible = False
        GoTo Fim
    Else
        flxArquivos.Visible = True
    End If
    
    '' Preparar Grid
    With flxArquivos
        .RowHeightMin = 400
        .Rows = objBCol.Count + 1
        .Cols = 5
        .TextMatrix(0, 0) = "Path Arquivo"
        .TextMatrix(0, 1) = "Nome Arquivo"
        .TextMatrix(0, 2) = "Data"
        .TextMatrix(0, 3) = ""
        .TextMatrix(0, 4) = ""
        .row = 0
        .ColWidth(0) = 0
        .ColWidth(1) = 5000
        .ColWidth(2) = 1420
        .ColWidth(3) = 500
        .ColWidth(4) = 0
    End With
    
    '' Dados
    For i = 1 To objBCol.Count
        With flxArquivos
            .TextMatrix(i, 0) = objBCol(i).PastaArquivo
            .TextMatrix(i, 1) = objBCol(i).NomeArquivo
            .TextMatrix(i, 2) = Format$(objBCol(i).DataAnexo, "dd-MM-yyyy HH:mm")  'ArquivoExtensao(objBCol(i).NomeArquivo)
            .TextMatrix(i, 4) = objBCol(i).NomeArquivoFisico
            .row = i
            .Col = 3
            '.ColWidth(.Col) = pictArqAnexo.Width
            ''pictArqAnexo.Visible = True
            Set .CellPicture = pictArqAnexo
            .Refresh
        End With
    Next i
    
Fim:
    Set objA = Nothing
    Set objB = Nothing
    
End Sub

Private Function ArquivoExtensao(NomeArq As String) As String
    Dim arr() As String ' array tipo arq
    Dim max As Integer
    Dim strExt As String
    strExt = "---"
    
    On Error GoTo Fim
    
    arr = Split(NomeArq, ".")
    max = UBound(arr)
    strExt = arr(max)
    
Fim:
    ArquivoExtensao = strExt
End Function

Private Sub Sair()
    If MsgBox("Parâmetros não informados." & vbCrLf & "Confirme fechar a tela?", vbYes + vbDefaultButton2, "Confirmar Fechar") = vbYes Then
        Unload Me
    End If
End Sub

Private Sub cmdCancelar_Click()
    Acao = Consultar
    Set ObjStc750B = Nothing
    
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
 '' Exclusão lógica dos selecionados
    Dim objA750 As New clsA46V750A
    Dim objA753 As New clsA46V753A
    
    Dim objB750 As New stcA46V750B
    Dim objB753 As New stcA46V753B
    Dim msg As String, i As Integer
    
    Set objA750.M46V999 = gM46V999
    Set objA753.M46V999 = gM46V999
'    objB750.IdObs = ObjStc750B.IdObs
'    objB750.IdSiscota = ObjStc750B.IdSiscota
'    objB750.Login = ObjStc750B.Login
'    objB750.Observacao = ObjStc750B.Observacao
'    objB750.Siscota = ObjStc750B.Siscota
'    objB750.Status = 1
    
    Set objB750 = ObjStc750B
    objB750.Observacao = txtObservacao.Text
    
    Set objB753 = ObjStc753B
    If objB753.ColecaoObs.Count > 0 Then
        For i = 1 To objB753.ColecaoObs.Count + 1
            If Not objB753.ColecaoObs(i) Is Nothing Then
                objB753.ColecaoObs.Remove (i)
            End If
        Next i
    End If
    
    If MsgBox("Deseja gravar a Observação?", vbYesNo + vbDefaultButton2, "Confirmar gravação") = vbYes Then
    
        objB753.ColecaoObs.Add objB750
    
        If (objA753.mfSalvar(objB753, gM46V999.gNomUsuario6, msg)) Then
            lblMensagem.Caption = "Executado com sucesso!"
            
            '' Refresh do Grid
            'Set ObjStc753B = objB753
            
            '' Refresh Gri
            Call frmT46V210A.CarregarObservacao
        Else
            lblMensagem.Caption = "Erro ao Excluir.  Retorno: " & msg
        End If
    
        cmdSalvar.Enabled = False
        txtObservacao.Enabled = False
    End If
    
    If Not objA750 Is Nothing Then
        Set objA750 = Nothing
    End If
    
    If Not objB750 Is Nothing Then
        Set objB750 = Nothing
    End If
End Sub

Private Sub flxArquivos_Click()
    '' Abrir arquivo selecionado
    If flxArquivos.row > 0 And flxArquivos.Col = 3 Then
        '' Abrir arquivo flxArquivos.TextMatrix(flxArquivos.row, 0)
        Dim myFSO As New FileSystemObject, myFile
        Dim Nome, pasta, pastaTemp As String
        
        lblMensagem.Caption = ""
        pastaTemp = App.Path & "\TEMP\"
        Nome = flxArquivos.TextMatrix(flxArquivos.row, 4)
        pasta = flxArquivos.TextMatrix(flxArquivos.row, 0)
        
        If Len(pasta) > 0 Then
            If Right(pasta, 1) <> "\" Then
                pasta = pasta & "\"
            End If
        End If
        
        If Not myFSO.FolderExists(pastaTemp) Then
            myFSO.CreateFolder (pastaTemp)
        End If
       
        If Not myFSO.FileExists(pasta & Nome) Then
            lblMensagem.Caption = "Arquivo indisponível no Servidor!"
        Else
            Set myFile = myFSO.GetFile(pasta & Nome)
        
            If Not myFile Is Nothing Then
                myFile.Copy (pastaTemp & Nome)
                lblMensagem.Caption = ArquivoAbrir(pastaTemp & Nome)
            Else
                lblMensagem.Caption = "Arquivo indisponível na máquina local."
            End If
        End If
    End If
End Sub

Private Function ArquivoAbrir(strNomeArq As String) As String
    Dim Mensagem As String
    
    On Error GoTo Erro:
    ShellExecute Me.hWnd, "open", strNomeArq, vbNullString, vbNullString, SW_SHOWNORMAL
    Mensagem = "Acessando " & strNomeArq
    GoTo Fim
    
Erro:
    Mensagem = "Erro: " & Err.Description
    
Fim:
    ArquivoAbrir = Mensagem
End Function

