VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT46V171A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selecione motivo de suspensão"
   ClientHeight    =   2460
   ClientLeft      =   4770
   ClientTop       =   2820
   ClientWidth     =   3720
   ControlBox      =   0   'False
   Icon            =   "T46V171A.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   Begin Threed.SSCommand cmdOk 
      Height          =   360
      Left            =   615
      TabIndex        =   6
      Top             =   2025
      Width           =   1065
      _Version        =   65536
      _ExtentX        =   1879
      _ExtentY        =   635
      _StockProps     =   78
      Caption         =   "&Ok"
      ForeColor       =   8388608
   End
   Begin VB.Frame fmMotivoSuspensao 
      Caption         =   "Motivos de Suspensão :"
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   1050
      Width           =   3570
      Begin VB.ComboBox cboMotivo 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   75
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   255
         Width           =   3330
      End
   End
   Begin VB.Frame fraOpcao 
      Caption         =   "Enviar fluxo para :"
      ForeColor       =   &H00800000&
      Height          =   840
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3570
      Begin VB.OptionButton optOpcao 
         Caption         =   "Resseguro"
         ForeColor       =   &H00800000&
         Height          =   400
         Index           =   2
         Left            =   2205
         TabIndex        =   3
         Top             =   240
         Width           =   1170
      End
      Begin VB.OptionButton optOpcao 
         Caption         =   "Grupo"
         ForeColor       =   &H00800000&
         Height          =   400
         Index           =   0
         Left            =   105
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   750
      End
      Begin VB.OptionButton optOpcao 
         Caption         =   "Produto"
         ForeColor       =   &H00800000&
         Height          =   405
         Index           =   1
         Left            =   1125
         TabIndex        =   1
         Top             =   240
         Width           =   960
      End
   End
   Begin Threed.SSCommand cmdCancelar 
      Height          =   360
      Left            =   2130
      TabIndex        =   7
      Top             =   2025
      Width           =   1065
      _Version        =   65536
      _ExtentX        =   1879
      _ExtentY        =   635
      _StockProps     =   78
      Caption         =   "&Cancelar"
      ForeColor       =   8388608
   End
End
Attribute VB_Name = "frmT46V171A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lStatusAcao As Integer
Dim bdGed As ADODB.Connection
Dim lMensagem As String

Private Sub cmdCancelar_Click()
    lStatusAcao = 0
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim lMotivo As Long
    Dim lFluxo As Integer        'Tipo de transmissão
    '1 = Grupo
    '2 = Produto
    '3 = Resseguro

    '1. Consiste opção.
    If optOpcao(0).Value = False And optOpcao(1).Value = False And optOpcao(2).Value = False Then
        Call gM46V999.gpGraLog(1, "FNP0017 - Opção não selecionada.")
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    If cboMotivo.Text = "" Then
        Call gM46V999.gpGraLog(1, "FNP0018 - Motivo não selecionado.")
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    If optOpcao(0).Value = True Then  'grupo 9164 ou 9167 ou 9170
        lFluxo = 2
        lStatusAcao = 40
    ElseIf optOpcao(1).Value = True Then    'produto 9165 ou 9168 ou 9171
        lFluxo = 3
        lStatusAcao = 10
    ElseIf optOpcao(2).Value = True Then    'resseguro 9166 ou 9169 ou 9172
        lFluxo = 9
        lStatusAcao = 35
    End If

    lMotivo = cboMotivo.ItemData(cboMotivo.ListIndex)

    If Not gM46V999.gfAbrBasDados("GED", gM46V999.gTipGed, gM46V999.gEndGed, bdGed, lMensagem) Then
        Call gM46V999.gpGraLog(1, lMensagem)
        Exit Sub
    End If

    If Not mfVerEnvioFluxo Then
        Call cmdCancelar_Click
        Call gM46V999.gpFecha2(bdGed)
        Exit Sub
    End If

    If Not mfSalvarTabPropMensagem(lFluxo, lMotivo) Then
        lStatusAcao = 0
        GoTo Finaliza
    End If

    If Not mfSalvarTabStatEmis(lStatusAcao) Then
        lStatusAcao = 0
        GoTo Finaliza
    End If

    If Not mfSalvarTabCtrlEmisGed(lStatusAcao) Then
        lStatusAcao = 0
        GoTo Finaliza
    End If


Finaliza:

    Call gM46V999.gpFecha2(bdGed)
    Unload Me

End Sub

Public Sub mpInstanciarForm(ByRef pAcao As Integer, Optional pHabilitarCancelar As Boolean = True)
'Habilita ou não o cancelar
    cmdCancelar.Enabled = pHabilitarCancelar

    Me.Show vbModal
    'lStatusAcao =  40 (grupo) ; 10 (produto) ; 35 (resseguro) ; 0 (nao suspender) ; 2(suspender)
    pAcao = lStatusAcao
    'pEnviarFluxo = lFluxo
    'pCodMotivo = lMotivo

End Sub

Private Sub Form_Activate()
'    frmT46V100A.panForm.Caption = Mid$(Me.Name, 4)
    If cmdOk.Visible = True And cmdOk.Enabled = True Then
        cmdOk.SetFocus
    End If
End Sub

Private Sub Form_Load()
'Centralizar form
    Call gM46V999.gpCenForm(Me)

    ' Preencher Combo de Motivo com Tabela TBS_TIPO
    optOpcao(0).Value = True
    Call mpCarregarMotivos(2)
    cboMotivo.ListIndex = 0

End Sub

Private Sub mpCarregarMotivos(pCodEvento As Integer)

    Dim lSelect As String
    Dim lrst As clsYasRecordSet
    Dim lMensagem As String
    Dim lAux As String
    Dim bdGed As ADODB.Connection


    If Not gM46V999.gfAbrBasDados("GED", gM46V999.gTipGed, gM46V999.gEndGed, bdGed, lMensagem) Then
        Call gM46V999.gpGraLog(1, lMensagem)
        Exit Sub
    End If

    lSelect = "Select * from ged.dbo.tbs_tipo where cod_tipo > 9163 and cod_tipo <= 9172 and "
    lSelect = lSelect & "cod_evento = " & pCodEvento

    If Not gM46V999.gfObtRegistro(bdGed, lSelect, lrst, lMensagem) Then
        Call gM46V999.gpFecha2(bdGed)
        Call gM46V999.gpGraLog(1, lMensagem)
        Exit Sub
    End If

    cboMotivo.Clear
    cboMotivo.AddItem ""

    '3. Carrega Combo
    Do While Not lrst.EOF
        cboMotivo.AddItem lrst("msg_recusa")
        cboMotivo.ItemData(cboMotivo.NewIndex) = lrst("Cod_Tipo")
        lrst.MoveNext
    Loop

    Set lrst = Nothing
    Call gM46V999.gpFecha2(bdGed)

    Exit Sub

End Sub
Private Function mfSalvarTabPropMensagem(pFluxo As Integer, pMotivo As Long) As Boolean

    Dim lSql As String
    Dim lAux As String

    mfSalvarTabPropMensagem = False

    lSql = "Delete from ged.dbo.Tab_Prop_Mensagem where num_pi  = '" & gM46V111.gNum_Pi & "' and "
    lSql = lSql & "cod_mensagem = " & pMotivo
    lSql = lSql & " and dat_digit = " & gM46V999.gHoje

    If gM46V999.gfExeSQL(bdGed, lSql, lMensagem) <> 0 Then
        Call gM46V999.gpGraLog(1, lMensagem)
        Exit Function
    End If

    lSql = " INSERT INTO ged.dbo.Tab_Prop_Mensagem values "
    lSql = lSql & "(" & Format(gM46V111.gNum_Pi, "0000000000") & ","
    lSql = lSql & " 0000, "
    lSql = lSql & pMotivo & ","
    lSql = lSql & pFluxo & ","
    lSql = lSql & "'',"
    lSql = lSql & gM46V999.gHoje & ",0,0)"

    If gM46V999.gfExeSQL(bdGed, lSql, lMensagem) <> 0 Then
        Call gM46V999.gpGraLog(1, lMensagem)
        Exit Function
    End If

    mfSalvarTabPropMensagem = True
    Exit Function

End Function

Private Function mfSalvarTabStatEmis(pStatusAcao As Integer) As Boolean

    Dim lSql As String
    Dim lStatusAcao As Integer
    Dim lDataAcao As Double

    mfSalvarTabStatEmis = False

    lSql = "Delete from ged.dbo.Tab_Stat_Emis where num_pi  = '" & gM46V111.gNum_Pi & "' and "
    lSql = lSql & "status_acao = " & pStatusAcao
    lSql = lSql & " and status_agente = 0"

    If gM46V999.gfExeSQL(bdGed, lSql, lMensagem) <> 0 Then
        Call gM46V999.gpGraLog(1, lMensagem)
        Exit Function
    End If

    '***marcia
    If gM46V999.gHojDia = 1 Then
        lDataAcao = Format(DateAdd("d", -1, gM46V999.gForHoje), "yyyymmdd")
    Else
        lDataAcao = gM46V999.gHoje - 1
    End If
    '***marcia

    lSql = "Delete from ged.dbo.Tab_Stat_Emis where num_pi  = '" & gM46V111.gNum_Pi & "' and "
    lSql = lSql & "status_acao = " & pStatusAcao
    'lSql = lSql & " and data_acao = " & gM46V999.gHoje - 1
    lSql = lSql & " and data_acao = " & lDataAcao

    If gM46V999.gfExeSQL(bdGed, lSql, lMensagem) <> 0 Then
        Call gM46V999.gpGraLog(1, lMensagem)
        Exit Function
    End If

    lSql = " INSERT INTO ged.dbo.tab_stat_emis values"
    lSql = lSql & "(" & gM46V111.gNum_Pi & ","
    lSql = lSql & pStatusAcao & ",0,0, "
    'lSql = lSql & gM46V999.gHoje - 1 & ")"
    lSql = lSql & lDataAcao & ")"


    If gM46V999.gfExeSQL(bdGed, lSql, lMensagem) <> 0 Then
        Call gM46V999.gpGraLog(1, lMensagem)
        Exit Function
    End If



    mfSalvarTabStatEmis = True
    Exit Function

End Function

Private Function mfSalvarTabCtrlEmisGed(pStatusAcao As Integer) As Boolean

    Dim lSql As String

    mfSalvarTabCtrlEmisGed = False

    lSql = " UPDATE ged.dbo.TAB_CTRL_EMIS SET "
    lSql = lSql & "cod_iden_stat_acao = " & pStatusAcao
    lSql = lSql & " WHERE num_pi = " & gM46V111.gNum_Pi

    If gM46V999.gfExeSQL(bdGed, lSql, lMensagem) <> 0 Then
        Call gM46V999.gpGraLog(1, lMensagem)
        Exit Function
    End If

    mfSalvarTabCtrlEmisGed = True
    Exit Function

End Function

Private Function mfVerEnvioFluxo() As Boolean

    Dim lSql As String
    Dim lStatusAcao As Integer
    Dim lrst As clsYasRecordSet
    Dim lEnviado As String

    mfVerEnvioFluxo = False

    lSql = "Select * from ged.dbo.Tab_Prop_Mensagem where num_pi  = '" & gM46V111.gNum_Pi & "'"
    lSql = lSql & " and cod_mensagem in (9164,9165,9166,9168,9169,9170,9171,9172)"
    lSql = lSql & " and dat_liber = 0"

    If Not gM46V999.gfObtRegistro(bdGed, lSql, lrst, lMensagem) Then
        Call gM46V999.gpGraLog(1, lMensagem)
        Exit Function
    End If

    If Not lrst.EOF Then
        If lrst("Cod_tip_mensagem") = 2 Then
            lEnviado = " o Grupo Análise."
        ElseIf lrst("Cod_tip_mensagem") = 3 Then
            lEnviado = " o Produto."
        ElseIf lrst("Cod_tip_mensagem") = 9 Then
            lEnviado = " o Resseguro."
        End If
        MsgBox "Não foi possível realizar a operação! O fluxo já foi encaminhado para " & lEnviado, vbInformation
        mfVerEnvioFluxo = False
        Exit Function
    End If

    mfVerEnvioFluxo = True
    Exit Function

End Function

Private Sub optOpcao_Click(Index As Integer)
    Dim lCodEvento As Integer

    If Index = 0 Then
        lCodEvento = 2
    ElseIf Index = 1 Then
        lCodEvento = 3
    ElseIf Index = 2 Then
        lCodEvento = 9
    End If
    Call mpCarregarMotivos(lCodEvento)
    cboMotivo.ListIndex = 0
End Sub
