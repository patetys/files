VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmT46V100A 
   Caption         =   "SYAS - Sistema Yasuda Seguros - Emissão"
   ClientHeight    =   7905
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "T46V100A.frx":0000
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7905
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cdgHelp 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Threed.SSPanel panRodape 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   7485
      Width           =   11880
      _Version        =   65536
      _ExtentX        =   20955
      _ExtentY        =   741
      _StockProps     =   15
      ForeColor       =   -2147483640
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
      Begin Threed.SSPanel panForm 
         Height          =   300
         Left            =   60
         TabIndex        =   1
         ToolTipText     =   "Nome do programa"
         Top             =   60
         Width           =   1000
         _Version        =   65536
         _ExtentX        =   1773
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "form"
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin Threed.SSPanel panData 
         Height          =   300
         Left            =   9600
         Negotiate       =   -1  'True
         TabIndex        =   6
         ToolTipText     =   "Data do dia."
         Top             =   60
         Width           =   1200
         _Version        =   65536
         _ExtentX        =   2117
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "dd/mm/aaaa"
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin Threed.SSPanel panBasDados 
         Height          =   300
         Left            =   1100
         TabIndex        =   2
         ToolTipText     =   "Localização das bases das bases de dados."
         Top             =   60
         Width           =   3700
         _Version        =   65536
         _ExtentX        =   6526
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin Threed.SSPanel panVersao 
         Height          =   300
         Left            =   4800
         TabIndex        =   3
         ToolTipText     =   "Versão do sistema."
         Top             =   60
         Width           =   1600
         _Version        =   65536
         _ExtentX        =   2822
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "Versão dd/mm/aaaa"
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin Threed.SSPanel panCorretor 
         Height          =   300
         Left            =   6400
         TabIndex        =   4
         ToolTipText     =   "Código do corretor."
         Top             =   60
         Width           =   1600
         _Version        =   65536
         _ExtentX        =   2822
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "Corretor XXXXXXX"
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin Threed.SSPanel panUsuario 
         Height          =   300
         Left            =   8000
         TabIndex        =   5
         ToolTipText     =   "Código do usuário."
         Top             =   60
         Width           =   1600
         _Version        =   65536
         _ExtentX        =   2822
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "Usuário XXXXXXXX"
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
   End
   Begin Threed.SSPanel panAviso 
      Height          =   1005
      Left            =   3510
      TabIndex        =   7
      Top             =   2970
      Visible         =   0   'False
      Width           =   4005
      _Version        =   65536
      _ExtentX        =   7056
      _ExtentY        =   1764
      _StockProps     =   15
      Caption         =   "Aguarde...   Atualizando dados do sistema..."
      ForeColor       =   128
      BackColor       =   -2147483644
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Outline         =   -1  'True
   End
End
Attribute VB_Name = "frmT46V100A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Activate()
    If gM46V111.gSaida = True Then
        Unload Me
    End If
End Sub

Sub mfVerificaVersaoAntigo()
    Dim lbdTAB_Produto As ADODB.Connection    'Base de dados .
    Dim lbdTAB_Produto2 As ADODB.Connection    'Base de dados .
    Dim lMensagem As String    'Mensagem.
    Dim lrsTAB_Produto As clsYasRecordSet    'Registro .
    Dim lrsTAB_Produto2 As clsYasRecordSet    'Registro .
    Dim NomePasta As String
    Dim lTotErr As Integer
    Dim lSelect As String


    On Error Resume Next

    If gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Or gM46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda Then
        'Or Not gM46V999.gBaseLocal Or _
         'gM46V999.gTipoHomologacao = "0" Then
        GoSub ExcluirBaseTemp
        Exit Sub
    End If

    If Dir("L:\") = "" Then
        NomePasta = "C:\Temp\Syas\"
        MkDir "C:\Temp\"
    Else
        NomePasta = "L:\User\Syas\"
        MkDir "L:\User\"
    End If
    lTotErr = 0

TentaNovamente:
    Err = 0
    If Dir(NomePasta) = "" Then
        MkDir NomePasta
        If Err <> 0 Then
            Call gM46V999.gpGraLog(0, "Não foi possível criar " & NomePasta & ", Erro: " & Err.Description)
            If lTotErr = 0 Then
                NomePasta = "C:\Temp\Syas\"
                lTotErr = 1
                GoTo TentaNovamente
            End If
            GoSub ExcluirBaseTemp
            Exit Sub
        End If
    End If

    If Dir(NomePasta & "P00MULTI.MDB") = "" Then
        GoTo Copia:
    End If

    If gM46V999.gfAbrBasDados("P00MULTI2", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTAB_Produto, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If

    If gM46V999.gfAbrBasDados("P00MULTI2", gM46V999.gTipP00MULTI, NomePasta & "P00MULTI.MDB", lbdTAB_Produto2, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If


    lSelect = "SELECT * FROM Tab_Produto "
    lSelect = lSelect & " order by Dat_Alt desc "

    If gM46V999.gfObtRegistro(lbdTAB_Produto, lSelect, lrsTAB_Produto, lMensagem) = False Then
        Call gM46V999.gpFecha2(lbdTAB_Produto)
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If

    If gM46V999.gfObtRegistro(lbdTAB_Produto2, lSelect, lrsTAB_Produto2, lMensagem) = False Then
        Call gM46V999.gpFecha2(lbdTAB_Produto)
        Call gM46V999.gpFecha2(lbdTAB_Produto2)
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If

    If Format(lrsTAB_Produto!Dat_Alt, "00000000000000") > Format(lrsTAB_Produto2!Dat_Alt, "00000000000000") Then
        GoTo Copia
    End If

    gM46V999.gEndP00MULTI = NomePasta & "P00MULTI.MDB"
    gM46V999.gTipoHomologacao = 1
    Exit Sub

Copia:
    panAviso.Caption = "Atualizando Banco Dados Local. Tempo de Espera:      5 minutos. Por favor aguarde."
    panAviso.Visible = True
    panAviso.Refresh

    Call gM46V999.gpFecha2(lbdTAB_Produto)
    Call gM46V999.gpFecha2(lbdTAB_Produto2)
    Err = 0
    FileCopy gM46V999.gEndP00MULTI, NomePasta & "P00MULTI.MDB"
    panAviso.Visible = False
    If Err <> 0 Then
        MsgBox Error
    End If
    If Dir(NomePasta & "P00MULTI.MDB") = "" Then
        Call gM46V999.gpFecha2(lbdTAB_Produto)
        Call gM46V999.gpFecha2(lbdTAB_Produto2)
        Call gM46V999.gpGraLog(2, "Não foi possível copia a base para a base CITRIX! Sistema Encerrado.")
        End
    End If
    gM46V999.gEndP00MULTI = NomePasta & "P00MULTI.MDB"
    gM46V999.gTipoHomologacao = 1
    Exit Sub

ExcluirBaseTemp:

    gM46V999.gTipoHomologacao = 0

    If Dir("L:\") = "" Then
        Return
    End If
    If Dir("L:\USER\SYAS\P00MULTI.MDB") <> "" Then
        On Error Resume Next
        Err = 0
        Kill "L:\USER\SYAS\P00MULTI.MDB"
        If Err <> 0 Then
            Call gM46V999.gpGraLog(2, "Não foi possível remover a base CITRIX! Sistema Encerrado.")
            End
        End If
    End If
    Return

End Sub


Sub mfVerificaVersao()
    Dim lbdTAB_Produto As ADODB.Connection    'Base de dados .
    Dim lbdTAB_Produto2 As ADODB.Connection    'Base de dados .
    Dim lMensagem As String    'Mensagem.
    Dim lrsTAB_Produto As clsYasRecordSet    'Registro .
    Dim lrsTAB_Produto2 As clsYasRecordSet    'Registro .
    Dim NomePasta As String
    Dim lSelect As String


    On Error Resume Next

    If gM46V999.gTipUsuario = e_Tipo_Usu_Corretor Or gM46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda Then
        GoSub ExcluirBaseTemp
        Exit Sub
    End If

    NomePasta = Empty
    If gfIsMetaframe() Then
        NomePasta = "L:\User\Syas\"
        MkDir "L:\User\"
        Call gM46V999.gpGraLog(0, "Eh Metaframe forcou o L " & NomePasta)
    Else
        GoSub ExcluirBaseTemp
        '    Comentado teste
        'RodrigoS
        NomePasta = "C:\Temp\Syas\"
        MkDir "C:\Temp\"
    End If

    MkDir NomePasta

    If Dir(NomePasta & "P00MULTI.MDB") = "" Then
        GoTo Copia:
    End If

    If gM46V999.gfAbrBasDados("P00MULTI2", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lbdTAB_Produto, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If

    If gM46V999.gfAbrBasDados("P00MULTI2", gM46V999.gTipP00MULTI, NomePasta & "P00MULTI.MDB", lbdTAB_Produto2, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If


    lSelect = "SELECT * FROM Tab_Produto "
    lSelect = lSelect & " order by Dat_Alt desc "

    If gM46V999.gfObtRegistro(lbdTAB_Produto, lSelect, lrsTAB_Produto, lMensagem) = False Then
        Call gM46V999.gpFecha2(lbdTAB_Produto)
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If

    If gM46V999.gfObtRegistro(lbdTAB_Produto2, lSelect, lrsTAB_Produto2, lMensagem) = False Then
        Call gM46V999.gpFecha2(lbdTAB_Produto)
        Call gM46V999.gpFecha2(lbdTAB_Produto2)
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If

    If Format(lrsTAB_Produto!Dat_Alt, "00000000000000") > Format(lrsTAB_Produto2!Dat_Alt, "00000000000000") Then
        GoTo Copia
    End If

    gM46V999.gEndP00MULTI = NomePasta & "P00MULTI.MDB"
    gM46V999.gTipoHomologacao = 1
    Exit Sub

Copia:
    panAviso.Caption = "Atualizando Banco Dados Local. Tempo de Espera:      5 minutos. Por favor aguarde."
    panAviso.Visible = True
    panAviso.Refresh

    Call gM46V999.gpFecha2(lbdTAB_Produto)
    Call gM46V999.gpFecha2(lbdTAB_Produto2)
    Err = 0
    FileCopy gM46V999.gEndP00MULTI, NomePasta & "P00MULTI.MDB"
    panAviso.Visible = False
    If Err <> 0 Then
        MsgBox Error
    End If
    If Dir(NomePasta & "P00MULTI.MDB") = "" Then
        Call gM46V999.gpFecha2(lbdTAB_Produto)
        Call gM46V999.gpFecha2(lbdTAB_Produto2)
        Call gM46V999.gpGraLog(2, "Não foi possível copia a base para a base CITRIX! Sistema Encerrado.")
        End
    End If
    gM46V999.gEndP00MULTI = NomePasta & "P00MULTI.MDB"
    gM46V999.gTipoHomologacao = 1
    Exit Sub

ExcluirBaseTemp:

    gM46V999.gTipoHomologacao = 0

    If Dir("L:\") = "" Then
        Return
    End If
    If Dir("L:\USER\SYAS\P00MULTI.MDB") <> "" Then
        On Error Resume Next
        Err = 0
        Kill "L:\USER\SYAS\P00MULTI.MDB"
        If Err <> 0 Then
            Call gM46V999.gpGraLog(2, "Não foi possível remover a base CITRIX! Sistema Encerrado.")
            End
        End If
    End If
    Return

End Sub




Private Sub panBasDados_DblClick()
    MsgBox "Endereço das bases de dados: " & panBasDados.Caption
End Sub
Private Sub panCorretor_DblClick()
    MsgBox "Código do corretor: " & Mid$(panCorretor.Caption, 10)
End Sub
Private Sub panData_DblClick()
    MsgBox "Data de hoje registrado neste microcomputador: " & panData.Caption
End Sub
Private Sub panForm_DblClick()
    MsgBox "Programa ativo: " & panForm.Caption
End Sub
Private Sub panUsuario_DblClick()
    MsgBox "Código do usuário: " & Mid(panUsuario.Caption, 9)
End Sub
Private Sub panVersao_DblClick()
    MsgBox "Data da versão do programa: " & Mid(panVersao.Caption, 8)
End Sub
