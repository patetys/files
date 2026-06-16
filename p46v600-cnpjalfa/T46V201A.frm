VERSION 5.00
Begin VB.Form frmT46V201A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Administração do Sistema"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7650
   Icon            =   "T46V201A.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame 
      Caption         =   "Regras"
      ForeColor       =   &H00800000&
      Height          =   795
      Left            =   90
      TabIndex        =   25
      Top             =   5160
      Width           =   7455
      Begin VB.CheckBox Check1 
         Caption         =   "Modo Fast Endosso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2400
         TabIndex        =   27
         Top             =   270
         Width           =   2025
      End
      Begin VB.CheckBox Check 
         Caption         =   "Desabilitar Normas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   26
         Top             =   270
         Width           =   1965
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Retransmissão Imobiliario (Funciona apenas para propostas so Imobiliario)"
      ForeColor       =   &H00800000&
      Height          =   975
      Left            =   60
      TabIndex        =   21
      Top             =   4110
      Width           =   7485
      Begin VB.TextBox txtNum_PI_Imobiliario 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1260
         MaxLength       =   10
         TabIndex        =   23
         Top             =   390
         Width           =   1065
      End
      Begin VB.CommandButton cmdImobiliario 
         Caption         =   "Ok"
         Height          =   350
         Left            =   2340
         TabIndex        =   22
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Num_PI:  "
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
         Left            =   150
         TabIndex        =   24
         Top             =   390
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Boletos"
      ForeColor       =   &H00800000&
      Height          =   795
      Left            =   60
      TabIndex        =   16
      Top             =   3240
      Width           =   7485
      Begin VB.CommandButton cmdDisponibilizarBoleto 
         Caption         =   "Disponibilizar"
         Height          =   350
         Left            =   120
         TabIndex        =   20
         Top             =   300
         Width           =   1110
      End
      Begin VB.TextBox txtQtdBoletoReservado 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2310
         MaxLength       =   10
         TabIndex        =   18
         Top             =   330
         Width           =   795
      End
      Begin VB.CommandButton cmdReservarBoleto 
         Caption         =   "Reservar"
         Height          =   350
         Left            =   3120
         TabIndex        =   17
         Top             =   300
         Width           =   900
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Qtd:  "
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
         Left            =   1680
         TabIndex        =   19
         Top             =   330
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Retransmissão"
      ForeColor       =   &H00800000&
      Height          =   975
      Left            =   60
      TabIndex        =   12
      Top             =   2190
      Width           =   7485
      Begin VB.CommandButton cmdTransmissao 
         Caption         =   "Ok"
         Height          =   350
         Left            =   2340
         TabIndex        =   14
         Top             =   300
         Width           =   900
      End
      Begin VB.TextBox txtNum_PI_Transmissao 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1260
         MaxLength       =   10
         TabIndex        =   13
         Top             =   330
         Width           =   1065
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Num_PI:  "
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
         Left            =   150
         TabIndex        =   15
         Top             =   330
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Copiar Cotação"
      ForeColor       =   &H00800000&
      Height          =   975
      Left            =   60
      TabIndex        =   6
      Top             =   1140
      Width           =   7485
      Begin VB.ComboBox cboOrigem 
         Height          =   315
         ItemData        =   "T46V201A.frx":0A02
         Left            =   1050
         List            =   "T46V201A.frx":0A0F
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   330
         Width           =   1155
      End
      Begin VB.CommandButton cmdCopiarOk 
         Caption         =   "Ok"
         Height          =   350
         Left            =   6120
         TabIndex        =   11
         Top             =   300
         Width           =   900
      End
      Begin VB.TextBox txtNosso_numero 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4170
         MaxLength       =   20
         TabIndex        =   10
         Top             =   330
         Width           =   1905
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Origem:   "
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
         Top             =   330
         Width           =   915
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nosso Numero:  "
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
         Left            =   2580
         TabIndex        =   9
         Top             =   330
         Width           =   1575
      End
   End
   Begin VB.Frame fraAlterarStatus 
      Caption         =   "Mover PI da fila do Usuário"
      ForeColor       =   &H00800000&
      Height          =   975
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   7485
      Begin VB.ComboBox cboCod_Iden_Status_Acao 
         Height          =   315
         ItemData        =   "T46V201A.frx":0A22
         Left            =   4200
         List            =   "T46V201A.frx":0A2C
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   330
         Width           =   1905
      End
      Begin VB.TextBox txtNum_PI 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1260
         MaxLength       =   10
         TabIndex        =   2
         Top             =   330
         Width           =   1065
      End
      Begin VB.CommandButton cmdStatusOk 
         Caption         =   "Ok"
         Height          =   350
         Left            =   6090
         TabIndex        =   5
         Top             =   300
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ação:   "
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
         Left            =   3270
         TabIndex        =   3
         Top             =   330
         Width           =   915
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Num_PI:  "
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
         Left            =   150
         TabIndex        =   1
         Top             =   330
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmT46V201A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check_Click()
    gDesabilitaNormas = (Check.value = Checked)
End Sub

Private Sub Check1_Click()
    gModoFastEndosso = (Check1.value = Checked)
End Sub

Private Sub cmdCopiarOk_Click()
    Dim lNosso_Numero As String
    Dim lMensagem As String
    Dim lScript As String
    
    On Error GoTo Erro
    
    lMensagem = Empty
    Screen.MousePointer = vbHourglass
    
    If Len(txtNosso_numero.Text) < 20 Or Not IsNumeric(txtNosso_numero.Text) Then
        MsgBox "Nosso Numero inválido!", vbCritical
        GoTo SAIDA
    End If
    
    If Not (cboOrigem.Text = "Kit" Or cboOrigem.Text = "Web" Or cboOrigem.Text = "Ftp") Then
        MsgBox "Origem inválida!", vbCritical
        GoTo SAIDA
    End If
    
    lNosso_Numero = txtNosso_numero.Text
    If cboOrigem.Text = "Web" Then
        lNosso_Numero = "w" & lNosso_Numero
    Else
        If cboOrigem.Text = "Ftp" Then
            lNosso_Numero = "f" & lNosso_Numero
        End If
    End If
    
    lScript = gM46V999.gfCopDev(lNosso_Numero, lMensagem)
    
    If lMensagem <> Empty Then
        MsgBox lMensagem, vbCritical
        GoTo SAIDA
    End If
    
    lMensagem = Empty
    Call gM46V999.gpGraScriptDev(lScript, lMensagem)
    
    If lMensagem <> Empty Then
        MsgBox lMensagem, vbCritical
        GoTo SAIDA
    End If
    
    
    MsgBox "Script de copia realizado com sucesso."
    Call Shell("notepad " & gM46V999.gPathScript, vbNormalFocus)
    GoTo SAIDA

Erro:
    MsgBox Err.Number & " - " & Err.Description
SAIDA:
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdDisponibilizarBoleto_Click()
    Dim lBanco As ADODB.Connection
    Dim lMensagem As String
    Dim lrs As clsYasRecordSet
    Dim lSelect As String
    Dim lRegAfetados As Long
    Dim lAC As Long

    If gM46V999.gfAbrBasDados("RAMOSDIVERSOS", gTipDIVERSOS, gEndDIVERSOS, lBanco, lMensagem) = False Then
        MsgBox lMensagem, vbCritical
        GoTo SAIDA
    End If
    
    'Thais mudou de 150 para 120 o prazo de reciclagem dos boletos
    lSelect = "insert into syas_emis..tab_boleto_disponivel"
    lSelect = lSelect & " SELECT"
    lSelect = lSelect & "   DISTINCT NUM_BOLETO, null"
    lSelect = lSelect & " FROM"
    lSelect = lSelect & "   SYAS_EMIS..tab_ctrl_pagantec A"
    lSelect = lSelect & " WHERE"
    lSelect = lSelect & "   A.dat_venc > 0"
    lSelect = lSelect & " AND A.dat_venc < convert(varchar,getdate()-150,112)"
    lSelect = lSelect & " AND A.num_boleto > 800000"
    lSelect = lSelect & " AND   NOT EXISTS  ("
    lSelect = lSelect & "               select num_boleto "
    lSelect = lSelect & "               from  syas_emis..tab_ctrl_pagantec B"
    lSelect = lSelect & "               where dat_venc = 0 and B.num_boleto = A.num_boleto)"
    lSelect = lSelect & " AND   NOT EXISTS  ("
    lSelect = lSelect & "               select num_boleto "
    lSelect = lSelect & "               from  syas_emis..tab_ctrl_pagantec B"
    lSelect = lSelect & "               where dat_venc > 0 and dat_venc > convert(varchar,getdate()-150,112) "
    lSelect = lSelect & "               and b.num_boleto = A.num_boleto)"
    lSelect = lSelect & " AND   NOT EXISTS  ("
    lSelect = lSelect & "               SELECT num_boleto"
    lSelect = lSelect & "               from syas_emis..tab_boleto_disponivel B"
    lSelect = lSelect & "               WHERE   B.Num_Boleto = A.NUM_BOLETO)"

    lRegAfetados = 0
    If gM46V999.gfExeSQL(lBanco, lSelect, lMensagem, lRegAfetados) > 0 Then
        MsgBox lMensagem, vbCritical
        GoTo SAIDA
    End If
    
    MsgBox lRegAfetados & " Boletos disponibilizados!"

SAIDA:
    GoSub FechaObj
    Exit Sub
FechaObj:
    Set lrs = Nothing
    Return
End Sub

Private Sub cmdReservarBoleto_Click()
    Dim lMensagem As String
    Dim lSelect As String
    Dim lRegAfetados As Long
    Dim i As Integer
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim lTotal As Integer
    
    On Error GoTo TrataErro:
    Screen.MousePointer = vbHourglass

    If Len(txtQtdBoletoReservado.Text) < 1 Or Not IsNumeric(txtQtdBoletoReservado.Text) Then
        MsgBox "Qtd Invalida!", vbCritical
        Exit Sub
    End If
    
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    cn.ConnectionTimeout = 30
    cn.CommandTimeout = 30
    cn.Open "Provider=SQLOLEDB.1;Initial Catalog=RamosDiversos;Data Source=192.168.1.97", "workflow", "yasworkflow"
    
    i = 0
    Do
        lSelect = " exec syas_emis.[dbo].proc_GetListaBoletoSyas 1601, 999999"
        If rs.State = adStateOpen Then
            rs.Close
        End If
        
        rs.Open lSelect, cn
        
        If rs.EOF Then Exit Do
        
        Do While Not rs.EOF
            lSelect = "delete from  SyasPortal..p0045200 where num_boleto = " & rs("num_boleto")
            cn.Execute lSelect
            lSelect = "insert into  SyasPortal..p0045200 (num_boleto, status) values  ( " & rs("num_boleto") & ",0)"
            cn.Execute lSelect
            rs.MoveNext
        Loop
        i = i + 1
        If i = lTotal Then Exit Do
    Loop
    
    MsgBox "Fim!"
    
Fim:
    If cn.State = adStateOpen Then
        cn.Close
    End If
    
    Set cn = Nothing
    
    If rs.State = adStateOpen Then
        rs.Close
    End If
    Set rs = Nothing

    Screen.MousePointer = vbDefault
    Exit Sub
    
TrataErro:
    MsgBox "[ERR] cmdExecutar_Click - " & Err.Number & ": " & Err.Description
    Resume Fim
End Sub

Private Sub cmdStatusOk_Click()
    Dim lBanco As ADODB.Connection
    Dim lMensagem As String
    Dim lrs As clsYasRecordSet
    Dim lSelect As String
    Dim lRegAfetados As Long

    If Len(txtNum_PI.Text) < 10 Or Not IsNumeric(txtNum_PI.Text) Then
        MsgBox "NUM PI inválido!", vbCritical
        GoTo SAIDA
    End If
    
    If Not IsNumeric(cboCod_Iden_Status_Acao.Text) Then
        MsgBox "Código Status inválido!", vbCritical
        GoTo SAIDA
    End If

    If gM46V999.gfAbrBasDados("RAMOSDIVERSOS", gTipDIVERSOS, gEndDIVERSOS, lBanco, lMensagem) = False Then
        MsgBox lMensagem, vbCritical
        GoTo SAIDA
    End If
    
    lSelect = "UPDATE RAMOSDIVERSOS.DBO.Tab_CTRL_EMIS SET Cod_Iden_Stat_Acao = " & cboCod_Iden_Status_Acao.Text
    lSelect = lSelect & " Where cod_ramo in (112, 113, 114) "
    lSelect = lSelect & " AND Num_Pi = " & txtNum_PI.Text
    
    lRegAfetados = 0
    If gM46V999.gfExeSQL(lBanco, lSelect, lMensagem, lRegAfetados) > 0 Then
        MsgBox lMensagem, vbCritical
        GoTo SAIDA
    End If
    
    MsgBox lRegAfetados & " registro(s) afetado(s)!"

SAIDA:
    GoSub FechaObj
    Exit Sub
FechaObj:
    Set lrs = Nothing
    Return
End Sub

Private Sub cmdImobiliario_Click()
    Dim lBanco As ADODB.Connection
    Dim lMensagem As String
    Dim lrs As clsYasRecordSet
    Dim lSelect As String
    Dim lRegAfetados As Long

    If Len(txtNum_PI_Imobiliario.Text) < 10 Or Not IsNumeric(txtNum_PI_Imobiliario.Text) Then
        MsgBox "NUM PI inválido!", vbCritical
        GoTo SAIDA
    End If
    
    If gM46V999.gfAbrBasDados("RAMOSDIVERSOS", gTipDIVERSOS, gEndDIVERSOS, lBanco, lMensagem) = False Then
        MsgBox lMensagem, vbCritical
        GoTo SAIDA
    End If
    
    lSelect = "SELECT Num_PI FROM Ftp_Tab_ped WHERE COD_ESTIP = 20 AND Num_Pi = " & txtNum_PI_Imobiliario.Text
    
    If Not gM46V999.gfObtRegistro(lBanco, lSelect, lrs, lMensagem) Then
        MsgBox lMensagem, vbCritical
        GoTo SAIDA
    End If
    
    If lrs.EOF Then
        MsgBox "Este PI não existe ou não é Imobiliario!!!!!! Porrrrraaa !!!!!!!!!", vbCritical
        GoTo SAIDA
    End If
    
    lSelect = "BEGIN "
    lSelect = lSelect & " SET NOCOUNT ON                                                                             "
    lSelect = lSelect & " DECLARE @MENSAGEM     varchar(max),                                                       "
    lSelect = lSelect & "       @NUM_PI         numeric(18),                                                        "
    lSelect = lSelect & "       @NOSSO_NUMERO   varchar(20)                                                         "
    lSelect = lSelect & " SET       @NUM_PI         = " & txtNum_PI_Imobiliario.Text & "; "
    lSelect = lSelect & " SELECT    @NOSSO_NUMERO = NOSSO_NUMERO FROM RamosDiversos..Tab_Ped WHERE NUM_PI = @NUM_PI;    "
    lSelect = lSelect & " DELETE FROM RamosDiversos..Tab_Erro_Transmissao WHERE NUM_PI = @NUM_PI;                    "
    lSelect = lSelect & " DELETE FROM [P0044800].[dbo].[Tab_ConsDadosExternos] WHERE num_pi = @NUM_PI;               "
    lSelect = lSelect & " IF @NOSSO_NUMERO IS NOT NULL                                                               "
    lSelect = lSelect & "   EXEC dbo.PROC_LIMPADADOSMULTI @NOSSO_NUMERO, @MENSAGEM OUTPUT;                          "
    lSelect = lSelect & " SELECT @MENSAGEM as Mensagem;                                                              "
    lSelect = lSelect & " SELECT @NOSSO_NUMERO = NOSSO_NUMERO FROM FTP_TAB_PED WHERE NUM_PI = @NUM_PI;               "
    lSelect = lSelect & " DELETE FROM TAB_CTRL_TRANSMISSAO WHERE NUM_PI = @NUM_PI;                                   "
    lSelect = lSelect & " EXEC proc_StartFluxoEmissaoTransmissaoMulti @NOSSO_NUMERO, @NUM_PI;                        "
    lSelect = lSelect & " END"
    
    lRegAfetados = 0
    If gM46V999.gfExeSQL(lBanco, lSelect, lMensagem, lRegAfetados) > 0 Then
        MsgBox lMensagem, vbCritical
        GoTo SAIDA
    End If

    
    MsgBox "Concluido com sucesso !", vbInformation

SAIDA:
    GoSub FechaObj
    Exit Sub
FechaObj:
    Set lrs = Nothing
    Return
End Sub

Private Sub cmdTransmissao_Click()
    Dim lBanco As ADODB.Connection
    Dim lMensagem As String
    Dim lrs As clsYasRecordSet
    Dim lSelect As String
    Dim lRegAfetados As Long

    If Len(txtNum_PI_Transmissao.Text) < 10 Or Not IsNumeric(txtNum_PI_Transmissao.Text) Then
        MsgBox "NUM PI inválido!", vbCritical
        GoTo SAIDA
    End If
    
    If gM46V999.gfAbrBasDados("RAMOSDIVERSOS", gTipDIVERSOS, gEndDIVERSOS, lBanco, lMensagem) = False Then
        MsgBox lMensagem, vbCritical
        GoTo SAIDA
    End If
    
    'Verificar Valor de cobertura zerado Tarifa 11.48

    lSelect = "SELECT COUNT(*) Qtd "
    lSelect = lSelect & " FROM Tab_ped_Cobert A INNER JOIN Tab_ped B ON A.Num_PI = B.Num_PI "
    lSelect = lSelect & " WHERE A.Val_PR = 0 AND COD_Dizer <= 1 AND A.NUM_PI = " & txtNum_PI_Transmissao.Text
    
    If gM46V999.gfObtRegistro(lBanco, lSelect, lrs, lMensagem) = False Then
        Call gM46V999.gpFecha2(lBanco)
        GoTo SAIDA
    End If
    
    If Not lrs.EOF Then
        If lrs!Qtd > 0 Then
            Call gM46V999.gpGraLog(1, "Não é possivel transmitir esta proposta!" & vbCrLf & "Valor de cobertura zerado, recalcule e tente novamente!")
            GoTo SAIDA
        End If
    End If
    
    lSelect = "UPDATE RAMOSDIVERSOS.DBO.Tab_CTRL_EMIS SET cod_iden_stat_acao = 0, dat_digit_mf = "
    lSelect = lSelect & "CAST(YEAR(GETDATE()) AS CHAR(4)) + RIGHT('00' + CAST(MONTH(GETDATE()) AS VARCHAR(2)), 2) + RIGHT('00' + CAST(DAY(GETDATE()) AS VARCHAR(2)), 2)"
    lSelect = lSelect & " Where cod_ramo in (112, 113, 114, 300, 620, 710) "
    lSelect = lSelect & " AND Num_Pi = " & txtNum_PI_Transmissao.Text
    
    lRegAfetados = 0
    If gM46V999.gfExeSQL(lBanco, lSelect, lMensagem, lRegAfetados) > 0 Then
        MsgBox lMensagem, vbCritical
        GoTo SAIDA
    End If
    
    lSelect = "UPDATE RAMOSDIVERSOS.DBO.Tab_PED SET COD_POSI_ATUAL = 1, COD_TIP_DOC = 2 "
    lSelect = lSelect & " Where cod_ramo in (112, 113, 114, 300, 620, 710) "
    lSelect = lSelect & " AND Num_Pi = " & txtNum_PI_Transmissao.Text
    
    lRegAfetados = 0
    If gM46V999.gfExeSQL(lBanco, lSelect, lMensagem, lRegAfetados) > 0 Then
        MsgBox lMensagem, vbCritical
        GoTo SAIDA
    End If
    
'    lSelect = "select * from [GED].[dbo].[tab_ctrl_emis] where num_pi = " & txtNum_PI_Transmissao.Text
'
'    lRegAfetados = 0
'    If gM46V999.gfExeSQL(lBanco, lSelect, lMensagem, lRegAfetados) > 0 Then
'        MsgBox lMensagem, vbCritical
'        GoTo SAIDA
'    End If
'
'    If lRegAfetados = -1 Then
'        lSelect = "INSERT INTO [GED].[dbo].[tab_ctrl_emis]([num_pi], [num_bloq], [cod_unid], [cod_ramo], [dat_ini_vig], [dat_ter_vig], [cod_corr], [per_corr], [cod_corr_2], [per_corr_2], [cod_corr_3], [per_corr_3], [nom_seg], "
'        lSelect = lSelect & " [qtd_anex], [num_apol_ant], [dsc_obs_prop], [tip_cotac], [num_cotac], [per_juros_neg], [cod_envi_doc], [dat_digit], [hor_digit], [cod_iden_stat_acao], [cod_user_digit], [dat_digit_mf], [hor_digit_mf], "
'        lSelect = lSelect & " [cod_user_digit_mf], [val_pr_tot], [cod_iden_classe], [cod_iden_cobr], [num_prop], [dat_canc_apol], [cod_tip_doc], [dat_rec_cia], [num_ref_corr], [dat_emis], [num_apol], [num_endo], [val_pr_emis], [val_pr_liq], "
'        lSelect = lSelect & " [per_tot_coss], [qtd_parc], [cod_iden_stat_resp], [cod_iden_sac], [cod_prod], [val_is], [per_desc_coml], [num_cgc_cpf], [cod_iden_pessoa], [cod_iden_pag_antec]) "
'        lSelect = lSelect & " select Num_PI, [num_bloq], [cod_unid], [cod_ramo], [dat_ini_vig], [dat_ter_vig], [cod_corr], [per_corr], [cod_corr_2], [per_corr_2], [cod_corr_3], [per_corr_3], [nom_seg], [qtd_anex], [num_apol_ant], "
'        lSelect = lSelect & " [dsc_obs_prop], [tip_cotac], [num_cotac], [per_juros_neg], [cod_envi_doc], [dat_digit], [hor_digit], [cod_iden_stat_acao], [cod_user_digit], [dat_digit_mf], [hor_digit_mf], [cod_user_digit_mf], "
'        lSelect = lSelect & " [val_pr_tot], [cod_iden_classe], [cod_iden_cobr], 0, [dat_canc_apol], [cod_tip_doc], convert(varchar,getdate(),112), [num_ref_corr], [dat_emis], [num_apol], [num_endo], [val_pr_emis], [val_pr_liq], "
'        lSelect = lSelect & " [per_tot_coss], [qtd_parc], [cod_iden_stat_resp], [cod_iden_sac], [cod_prod], [val_is], [per_desc_coml], [num_cgc_cpf], [cod_iden_pessoa], [cod_iden_pag_antec] from ramosdiversos.dbo.tab_ctrl_emis where Num_Pi = " & txtNum_PI_Transmissao.Text
'
'        lRegAfetados = 0
'        If gM46V999.gfExeSQL(lBanco, lSelect, lMensagem, lRegAfetados) > 0 Then
'            MsgBox lMensagem, vbCritical
'            GoTo SAIDA
'        End If
'    End If
    
    MsgBox lRegAfetados & " registro(s) afetado(s)!"

SAIDA:
    GoSub FechaObj
    Exit Sub
FechaObj:
    Set lrs = Nothing
    Return
End Sub



'Private Sub Form_Load()
'    Dim teste As New clsA99V001
'    Dim dig As String
'
'    Call teste.DIG_BLOQUETO("058601266", dig, "")
'End Sub
Private Sub Form_Load()

End Sub
