VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmT46V210A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "     Observação"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11835
   ControlBox      =   0   'False
   Icon            =   "T46V210A.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   11835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid flxObservacao 
      Height          =   4215
      Left            =   120
      TabIndex        =   5
      Top             =   1500
      Width           =   11505
      _ExtentX        =   20294
      _ExtentY        =   7435
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      RowHeightMin    =   120
      BackColor       =   -2147483644
      ForeColor       =   8388608
      BackColorFixed  =   -2147483636
      ForeColorFixed  =   16777215
      BackColorSel    =   65535
      ForeColorSel    =   12648447
      BackColorBkg    =   -2147483633
      GridColorFixed  =   4210752
      AllowBigSelection=   0   'False
      GridLines       =   2
      AllowUserResizing=   1
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   120
      Width           =   720
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   120
      Width           =   720
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "Editar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   720
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "Incluir"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   720
   End
   Begin VB.Label lblMensagem 
      Caption         =   "Selecione item para consulta:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   4
      Top             =   1000
      Width           =   6000
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   11640
      Y1              =   840
      Y2              =   840
   End
End
Attribute VB_Name = "frmT46V210A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ConexaoBD As ADODB.Connection    ' Conexão com o Banco de Dados
Public UltimoErro As String             ' Descrição do último erro gerado

'' Variaveis da selecao local
Private vLogin As String
Private vNossoNumeroObs As String
Private vNumPed As Long
Private vSiscota As String
Private vObservacao As String
Private vIdGrupo As String
Private vIdObs As Long
Private vIdSiscota As Long


Public Sub Form_Load()
    vLogin = gM46V999.gNomUsuario6
    vNossoNumeroObs = gM46V111.gComNosNumero
    vNumPed = gM46V111.gNumPedido
    vSiscota = frmT46V102A.txtNumSiscota.Text
    vIdGrupo = ""
    vIdSiscota = 0
    
    Call CarregarObservacao
    
    Call gM46V999.gpCenForm(Me)
End Sub

Public Function FormIsLoaded(formName As String) As Boolean
    Dim frm As Form
    
    For Each frm In Forms
        If UCase(frm.Name) = UCase(formName) Then
            FormIsLoaded = True
            Exit For
        End If
    Next frm
End Function

Public Sub CarregarObservacao()
    Dim i As Integer, msg As String
    Dim objA As New clsA46V753A
    Dim objB As New Collection

    '' Zerar variaveis e botoes
    lblMensagem.Caption = "Selecione item para consulta:"
    cmdEditar.Enabled = False
    cmdExcluir.Enabled = False
    
    If gM46V999.gDesabilitarForms Then
        cmdIncluir.Enabled = False
    End If
        
    '' Conexao e variaveis globais
    Set objA.M46V999 = gM46V999

    '' Verificar se NossoNumero foi informado
    If gM46V111 Is Nothing Then
        Call DadosIndisponiveis
        GoTo Fim
    End If
    
    '' Carregar GRID
    If (objA.mfCarregarObservacoes(objB, vNossoNumeroObs, msg)) = False Then
        '' Não há observações, exibir msg e sair
        Call DadosIndisponiveis
        GoTo Fim
    End If
    
    If (objB Is Nothing Or objB.Count <= 0) Then
        '' Não há observações para o NossoNumero
        Call DadosIndisponiveis
        GoTo Fim
    End If
    
    '' Preparar Grid
    With flxObservacao
        .RowHeightMin = 360
        .Rows = objB.Count + 1
        .Cols = 6
        .TextMatrix(0, 0) = "Data"
        .TextMatrix(0, 1) = "Login"
        .TextMatrix(0, 2) = "Observação"
        .TextMatrix(0, 3) = "Referência"
        .TextMatrix(0, 4) = ""
        .TextMatrix(0, 5) = ""
        .row = 0
        .ColWidth(0) = 1600
        .ColWidth(1) = 1400
        .ColWidth(2) = 5000
        .ColWidth(3) = 1000
        .ColWidth(4) = 0
        .ColWidth(5) = 0
    End With
    
    '' Guardar collection e variavel de Grupo
    Set gM46V999.gStcA46V750BrptCollection = objB
    
    vIdSiscota = objB(1).IdSiscota
    vIdGrupo = objB(1).IdGrupo

    '' Carregar grid
    For i = 1 To objB.Count
        With flxObservacao
            .Col = 0
            .CellAlignment = flexAlignCenterCenter
            .TextMatrix(i, 0) = Format$(objB.Item(i).Data, "dd-MM-yyyy HH:mm")
            
            .Col = 1
            .CellAlignment = flexAlignCenterCenter
            .TextMatrix(i, 1) = objB.Item(i).Login
            
            .Col = 2
            .CellAlignment = flexAlignCenterCenter
            .TextMatrix(i, 2) = objB.Item(i).Observacao
            
            .Col = 3
            .CellAlignment = flexAlignCenterCenter
            .TextMatrix(i, 3) = objB.Item(i).Num_Ped
            
            .TextMatrix(i, 4) = objB.Item(i).IdObs
            .TextMatrix(i, 5) = objB.Item(i).IdSiscota
        End With
    Next
    
    '' Limpar checkbox
    Call FlexGridBackColor(0, False)
    
Fim:
    If Not objA Is Nothing Then
        Set objA = Nothing
    End If

    If Not objB Is Nothing Then
        Set objB = Nothing
    End If
End Sub

Public Sub flxObservacao_SelChange()
    Dim c As Integer
    
    '' Nova Selecao
    lblMensagem.Width = 8000
    lblMensagem.Caption = "Selecione item para consulta:" ''& flxObservacao.TextMatrix((flxObservacao.row), 4)
    
    '' Habilitar botoes
    cmdEditar.Enabled = True
    cmdExcluir.Enabled = True
    
    '' Selecionado
    Set gM46V999.gStcA46V750BrptSelected = gM46V999.gStcA46V750BrptCollection.Item(flxObservacao.row)
    
    Call FlexGridBackColor(flxObservacao.row, True)
    
    Call VerificarBotoes
End Sub

Private Sub FlexGridBackColor(row As Integer, STATUS As Boolean)
    Dim r, c As Integer
        
    For r = 1 To flxObservacao.Rows - 1
        For c = 0 To flxObservacao.Cols - 1
            With flxObservacao
                .row = r
                .Col = c
                If (STATUS And row = r) Then
                    .CellBackColor = &HFFFF&        '' Selecionado
                Else
                    .CellBackColor = &H8000000F     '' Limpar
                End If
            End With
        Next c
    Next r
End Sub

Private Sub VerificarBotoes()
    If gM46V999.gStcA46V750BrptSelected.CodOrigem = 2 Then
        cmdEditar.Enabled = True
        cmdExcluir.Enabled = True
    Else
        cmdEditar.Enabled = False
        cmdExcluir.Enabled = False
    End If
End Sub

Private Sub DadosIndisponiveis()
        lblMensagem.Caption = "Nenhuma informação de 'Observação' disponível."
        flxObservacao.Visible = False
        cmdEditar.Enabled = False
        cmdExcluir.Enabled = False
End Sub



Private Sub flxObservacao_DblClick()
    Dim frmT46211 As New frmT46V211A
    
    '' Necessario traduzir as classes 750B/753B e 750BRpt
    Dim objSisB As New stcA46V753B
    objSisB.IdSiscota = gM46V999.gStcA46V750BrptSelected.IdSiscota
    objSisB.IdGrupo = gM46V999.gStcA46V750BrptSelected.IdGrupo
    objSisB.NOSSO_NUMERO = gM46V999.gStcA46V750BrptSelected.NOSSO_NUMERO
    objSisB.Num_Ped = gM46V999.gStcA46V750BrptSelected.Num_Ped
    objSisB.Siscota = gM46V999.gStcA46V750BrptSelected.Siscota
    objSisB.STATUS = gM46V999.gStcA46V750BrptSelected.STATUS
    
    Dim objObsB As New stcA46V750B
    objObsB.IdObs = gM46V999.gStcA46V750BrptSelected.IdObs
    objObsB.IdSiscota = gM46V999.gStcA46V750BrptSelected.IdSiscota
    objObsB.Login = gM46V999.gStcA46V750BrptSelected.Login
    objObsB.CodOrigem = gM46V999.gStcA46V750BrptSelected.CodOrigem
    objObsB.Observacao = gM46V999.gStcA46V750BrptSelected.Observacao
    objObsB.STATUS = 1 ' gM46V999.gStcA46V750BrptSelected.StatusSiscota
    
    objSisB.ColecaoObs.Add objObsB
    
    Set frmT46211.ObjStc750B = objObsB
    Set frmT46211.ObjStc753B = objSisB
  
    frmT46211.Acao = Consultar
    frmT46211.Show 1
End Sub


Private Sub cmdIncluir_Click()
    Dim frmT46211 As New frmT46V211A
'
''    If Not FormIsLoaded(frmT46211.Name) Then
'    Dim objB As New stcA46V750B
'    objB.IdObs = 0
'    objB.IdSiscota = 0
'    objB.NOSSO_NUMERO = vNossoNumeroObs
'    objB.Num_Ped = vNumPed
'    objB.Observacao = ""
'    objB.Siscota = vSiscota
'    objB.Status = eStatus_Ativo
'
'    Set frmT46211.ObjStc750B = objB

    Dim objBSis As New stcA46V753B, objBObs As New stcA46V750B
    objBSis.NOSSO_NUMERO = vNossoNumeroObs
    objBSis.Num_Ped = vNumPed
    objBSis.Siscota = vSiscota
    objBSis.STATUS = 2
    objBSis.IdGrupo = vIdGrupo
    objBSis.IdSiscota = vIdSiscota
        
    objBObs.IdSiscota = objBSis.IdSiscota
    objBObs.IdObs = 0
    objBObs.CodOrigem = 2
    objBObs.Login = vLogin
    objBObs.Observacao = ""
    objBObs.STATUS = 1
        
    Set frmT46211.ObjStc753B = objBSis
    Set frmT46211.ObjStc750B = objBObs
        
    frmT46211.Acao = Incluir
    frmT46211.Show 1
        
'    Else
'        MsgBox ("Existe uma tela aberta para esta ação. Feche para prosseguir!")
'    End If
  
End Sub


Private Sub cmdEditar_Click()
    Dim frmT46211 As New frmT46V211A
    
    'If Not FormIsLoaded(frmT46211.Name) Then
    '' Necessario traduzir as classes 750B/753B e 750BRpt
    Dim objSisB As New stcA46V753B
    objSisB.IdSiscota = gM46V999.gStcA46V750BrptSelected.IdSiscota
    objSisB.IdGrupo = gM46V999.gStcA46V750BrptSelected.IdGrupo
    objSisB.NOSSO_NUMERO = gM46V999.gStcA46V750BrptSelected.NOSSO_NUMERO
    objSisB.Num_Ped = gM46V999.gStcA46V750BrptSelected.Num_Ped
    objSisB.Siscota = gM46V999.gStcA46V750BrptSelected.Siscota
    objSisB.STATUS = gM46V999.gStcA46V750BrptSelected.STATUS
    
    Dim objObsB As New stcA46V750B
    objObsB.IdObs = gM46V999.gStcA46V750BrptSelected.IdObs
    objObsB.IdSiscota = gM46V999.gStcA46V750BrptSelected.IdSiscota
    objObsB.Login = gM46V999.gStcA46V750BrptSelected.Login
    objObsB.CodOrigem = gM46V999.gStcA46V750BrptSelected.CodOrigem
    objObsB.Observacao = gM46V999.gStcA46V750BrptSelected.Observacao
    objObsB.STATUS = 1  'gM46V999.gStcA46V750BrptSelected.StatusSiscota
    
    Set frmT46211.ObjStc750B = objObsB
    Set frmT46211.ObjStc753B = objSisB
    
    frmT46211.Acao = Editar
    frmT46211.Show 1
        
'    Else
'        MsgBox ("Existe uma tela aberta para esta ação. Feche para prosseguir!")
'    End If
End Sub


Private Sub cmdExcluir_Click()
    '' Exclusão lógica dos selecionados
    Dim objA As New clsA46V750A
    Dim objB As New stcA46V750B
    Dim msg As String
    
    Set objA.M46V999 = gM46V999
    objB.IdObs = gM46V999.gStcA46V750BrptSelected.IdObs
    objB.IdSiscota = gM46V999.gStcA46V750BrptSelected.IdSiscota
    objB.STATUS = 2
        
    If MsgBox("Confirma a Exclusão da Observação selecionada?", vbYesNo + vbDefaultButton2, "Confirmar exclusão") = vbYes Then
        If (objA.mfExcluir(objB, NomeUsuarioLogin, msg)) Then
            lblMensagem.Caption = "Executado com sucesso!"
            
            '' Refresh do Grid
            Call CarregarObservacao
        Else
            lblMensagem.Caption = "Erro ao Excluir.  Retorno: " & msg
        End If
    End If
    
Fim:
    If Not objA Is Nothing Then
        Set objA = Nothing
    End If
    
    If Not objB Is Nothing Then
        Set objB = Nothing
    End If
End Sub


Private Sub cmdSair_Click()
    ''Check Forms relacionados e Dispose
    
    ''Close
    If Not gM46V999.gStcA46V750BrptCollection Is Nothing Then
        Set gM46V999.gStcA46V750BrptCollection = Nothing
    End If
    
    If Not gM46V999.gStcA46V750BrptSelected Is Nothing Then
        Set gM46V999.gStcA46V750BrptSelected = Nothing
    End If
    
    Unload Me
End Sub
