VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E7BC34A0-BA86-11CF-84B1-CBC2DA68BF6C}#1.0#0"; "NTSVC.ocx"
Begin VB.Form frmT46V157A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dados do Fluxo"
   ClientHeight    =   5550
   ClientLeft      =   1230
   ClientTop       =   1935
   ClientWidth     =   8955
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "T46V157A.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   Begin NTService.NTService NTService1 
      Left            =   7515
      Top             =   4770
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      ServiceName     =   "Simple"
      StartMode       =   3
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   945
      Left            =   2025
      MaxLength       =   30
      TabIndex        =   29
      ToolTipText     =   "Região de risco."
      Top             =   4590
      Width           =   4890
   End
   Begin VB.TextBox txtRisBairro 
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   6525
      MaxLength       =   30
      TabIndex        =   16
      ToolTipText     =   "Bairro do local de risco."
      Top             =   405
      Width           =   2310
   End
   Begin VB.TextBox txtRisNomLogradouro 
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   6525
      MaxLength       =   50
      TabIndex        =   15
      ToolTipText     =   "Nome do logradouro do endereço do segurado."
      Top             =   90
      Width           =   2310
   End
   Begin VB.TextBox txtRisCEP 
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1620
      MaxLength       =   9
      TabIndex        =   14
      ToolTipText     =   "CEP de risco. (xxxxx-xxx)"
      Top             =   90
      Width           =   2985
   End
   Begin VB.TextBox txtRisComplemento 
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1620
      MaxLength       =   20
      TabIndex        =   13
      ToolTipText     =   "Complemento."
      Top             =   405
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   1170
      Left            =   4815
      MaxLength       =   30
      TabIndex        =   10
      ToolTipText     =   "Região de risco."
      Top             =   1395
      Width           =   4080
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   1170
      Left            =   45
      MaxLength       =   30
      TabIndex        =   8
      ToolTipText     =   "Região de risco."
      Top             =   1395
      Width           =   4710
   End
   Begin Threed.SSCommand cmdResponder 
      Height          =   420
      Left            =   7785
      TabIndex        =   4
      Top             =   2610
      Width           =   1140
      _Version        =   65536
      _ExtentX        =   2011
      _ExtentY        =   741
      _StockProps     =   78
      Caption         =   "Responder"
      ForeColor       =   8388608
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1140
      Left            =   45
      TabIndex        =   3
      Top             =   2925
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   2011
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox txtRisRegiao 
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   1170
      Left            =   2880
      MaxLength       =   30
      TabIndex        =   0
      ToolTipText     =   "Região de risco."
      Top             =   2925
      Width           =   4890
   End
   Begin Threed.SSCommand cmdConsultar 
      Height          =   330
      Left            =   7785
      TabIndex        =   5
      Top             =   3780
      Width           =   1140
      _Version        =   65536
      _ExtentX        =   2011
      _ExtentY        =   582
      _StockProps     =   78
      Caption         =   "Consultar"
      ForeColor       =   8388608
   End
   Begin Threed.SSCommand cmdVerFluxo 
      Height          =   330
      Left            =   7785
      TabIndex        =   6
      Top             =   3465
      Width           =   1140
      _Version        =   65536
      _ExtentX        =   2011
      _ExtentY        =   582
      _StockProps     =   78
      Caption         =   "Ver Fluxo"
      ForeColor       =   8388608
   End
   Begin Threed.SSCommand cmdDelegar 
      Height          =   375
      Left            =   7785
      TabIndex        =   7
      Top             =   3105
      Width           =   1140
      _Version        =   65536
      _ExtentX        =   2011
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "Delegar"
      ForeColor       =   8388608
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   285
      Left            =   3465
      TabIndex        =   12
      Top             =   1080
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   503
      _StockProps     =   78
      Caption         =   "Enviar Email"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCommand cmdCepRisco 
      Height          =   285
      Left            =   1620
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   720
      Width           =   285
      _Version        =   65536
      _ExtentX        =   503
      _ExtentY        =   503
      _StockProps     =   78
      ForeColor       =   8388608
      Picture         =   "T46V157A.frx":030A
   End
   Begin Threed.SSCommand cmdFechar 
      Height          =   330
      Left            =   7785
      TabIndex        =   23
      Top             =   4140
      Width           =   1140
      _Version        =   65536
      _ExtentX        =   2011
      _ExtentY        =   582
      _StockProps     =   78
      Caption         =   "Fechar"
      ForeColor       =   0
   End
   Begin Threed.SSCommand cmdHistoricos 
      Height          =   330
      Left            =   45
      TabIndex        =   24
      Top             =   4140
      Width           =   2265
      _Version        =   65536
      _ExtentX        =   3995
      _ExtentY        =   582
      _StockProps     =   78
      Caption         =   "Histórico de Comentários"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCommand SSCommand5 
      Height          =   330
      Left            =   5985
      TabIndex        =   25
      Top             =   4140
      Width           =   1365
      _Version        =   65536
      _ExtentX        =   2408
      _ExtentY        =   582
      _StockProps     =   78
      Caption         =   "Consistências"
      ForeColor       =   0
   End
   Begin Threed.SSCommand cmdProduto 
      Height          =   330
      Left            =   4770
      TabIndex        =   26
      Top             =   4140
      Width           =   1185
      _Version        =   65536
      _ExtentX        =   2090
      _ExtentY        =   582
      _StockProps     =   78
      Caption         =   "Produto"
      ForeColor       =   0
   End
   Begin Threed.SSCommand cmdDigitacao 
      Height          =   330
      Left            =   3600
      TabIndex        =   27
      Top             =   4140
      Width           =   1140
      _Version        =   65536
      _ExtentX        =   2011
      _ExtentY        =   582
      _StockProps     =   78
      Caption         =   "Digitação"
      ForeColor       =   0
   End
   Begin Threed.SSCommand cmdTriagem 
      Height          =   330
      Left            =   2340
      TabIndex        =   28
      Top             =   4140
      Width           =   1230
      _Version        =   65536
      _ExtentX        =   2170
      _ExtentY        =   582
      _StockProps     =   78
      Caption         =   "Triagem"
      ForeColor       =   0
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mensagem da Carta de Recusa"
      ForeColor       =   &H00800000&
      Height          =   540
      Left            =   45
      TabIndex        =   30
      ToolTipText     =   "Unidade da federação do endereço do segurado."
      Top             =   4590
      Width           =   1905
   End
   Begin VB.Label lblMulti 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Processo:"
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   20
      Left            =   45
      TabIndex        =   22
      Top             =   405
      Width           =   1545
   End
   Begin VB.Label lblMulti 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Prazo Final do Passo:"
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   11
      Left            =   4680
      TabIndex        =   21
      ToolTipText     =   "Logradouro do endereço do segurado."
      Top             =   90
      Width           =   1815
   End
   Begin VB.Label lblMulti 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Data do Processo:"
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   10
      Left            =   45
      TabIndex        =   20
      Top             =   90
      Width           =   1545
   End
   Begin VB.Label lblMulti 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Assunto:"
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   12
      Left            =   4680
      TabIndex        =   19
      Top             =   405
      Width           =   1815
   End
   Begin VB.Label lblMulti 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Anexos:"
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   14
      Left            =   45
      TabIndex        =   18
      Top             =   720
      Width           =   1545
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Problemas Resolvidos"
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   4770
      TabIndex        =   11
      ToolTipText     =   "Unidade da federação do endereço do segurado."
      Top             =   1080
      Width           =   4110
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Problemas Pendentes"
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   45
      TabIndex        =   9
      ToolTipText     =   "Unidade da federação do endereço do segurado."
      Top             =   1080
      Width           =   3390
   End
   Begin VB.Label lblMulti 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Selecione a Resposta"
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   13
      Left            =   45
      TabIndex        =   2
      Top             =   2610
      Width           =   2760
   End
   Begin VB.Label lblRisRegiao 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Comentários/Justificativa"
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2880
      TabIndex        =   1
      ToolTipText     =   "Unidade da federação do endereço do segurado."
      Top             =   2610
      Width           =   4875
   End
End
Attribute VB_Name = "frmT46V157A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Dim ErroaoAbrir  As Boolean
Private Sub cmdFechar_Click()
    Me.Hide
''    Unload Me
End Sub

Private Sub Form_Load()
Dim ErrDescription As String
Dim tmpTipo        As String

    frmT46V100A.panForm.Caption = Mid$(Me.Name, 4)
    Call gpCenForm(Me)
    
    Dim objTaskMessage As KWOleSvr.TaskMessage
    Dim objFolder As MAPI.Folder
    Dim objMsgColl As MAPI.Messages
    Dim objMsgFilter As MAPI.MessageFilter
   
    On Error Resume Next
    Set objTaskMessage = CreateObject("Keyflow.TaskMessage.Document")
   
    If Err <> 0 Then
        Call gpGraLog(0, "[MSG] - Erro ao criar objeto objTaskMessage. Descrição: " & Err.Description)
        ErroaoAbrir = True
        Exit Sub
   End If
    
    If Not open_conect_mapi Then
        Call NTService1.LogEvent(svcEventInformation, svcMessageInfo, "Erro ao tentar efetuar Log On na MAPI.Session.")
        Exit Sub
    End If
    
    
    Set objFolder = objSession.Inbox
   
    If Err <> 0 Then
        ErrDescription = Err.Description
        Call gpGraLog(0, "[ERR] - Falha no objeto [ Set objFolder = objSession.Inbox ]. Descrição: " & ErrDescription)
        ErroaoAbrir = True
        Exit Sub
    End If
   '
    
   '
   '
            Set objMsgColl = objFolder.Messages
            Set objMsgFilter = objMsgColl.Filter
   '
c:
            tmpTipo = ""
            objMsgFilter.Subject = ""
'

            objMsgFilter.Subject = Format(gNum_Pi, "0000000000") & "-Triagem"
             tmpTipo = objMsgFilter.Subject
           
'
            If objMsgColl.Count = 0 Then
'                Call GeraLogs(CaminhoLogAgenteStartFlow, "[MSG] - Não encontrou o Passo: " & tmpTipo & " para ser processado.", lstLOG, True)
'                db.RollbackTrans
'                GoTo ProximoRegistro
            Else
               If objMsgColl.Count > 1 Then
                   objMsgColl.Item(1).Delete
                    If Err.Number <> 0 Then
'                         Call GeraLogs(CaminhoLogAgenteStartFlow, "[MSG] - Passo não respondido. Erro no Item.Delete(PassoEspec). Passo: " & Trim(rs("num_pi")), lstLOG, True)
'                         db.RollbackTrans
'                         GoTo ProximoRegistro
                    End If
                    GoTo c
               End If
            End If

'                Call GeraLogs(CaminhoLogAgenteStartFlow, "[MSG] - Processando Passo: " & tmpTipo & " ...", lstLOG, True)
'                Err.Clear
'                On Error Resume Next
'                Set objTaskMessage.SessionObject = objSession
'                objTaskMessage.EntryID = objMsgColl.Item(1).ID
'
 '               If Err <> 0 Then
'                    ErrDescription = Err.Description
'                    Call GeraLogs(CaminhoLogAgenteStartFlow, "[ERR] - Erro ao instanciar o fluxo, na rotina PassoEspecificacao(). Descrição: " & ErrDescription, lstLOG, True)
'                    Err.Clear
'                    GoTo ProximoRegistro
 '               End If
'
'                Call LiberaPITerminal(intRet, rs("num_pi"), rs2("num_prop"), rs2("cod_ramo"))
'
'                objTaskMessage.SelectedResponse = ""
'
'                If intRet = 0 Then
 ''                   objTaskMessage.SelectedResponse = "OK"
 '               End If
 '
 '               If objTaskMessage.SelectedResponse <> "" Then
 '                   SQL = ""
 '                   SQL = "Delete tab_prazo_vistoria where num_pi = " & Trim(rs("num_pi"))
 '                   SQL = SQL & " and acao = 65"
 '
 '                   Err.Clear
 '                   db.Execute SQL
 '                   If Err <> 0 Then
  '                      Call GeraLogs(CaminhoLogAgenteStartFlow, "[ERR] - Erro ao processar Passo: " & tmpTipo & ". SQL:" & SQL, lstLOG, True)
  '                      db.RollbackTrans
  ''                      GoTo ProximoRegistro
  '                  End If
  '
  '                  SQL = "Update TAB_STAT_EMIS Set status_resposta=1 where num_pi=" & rs("num_pi") & " and status_acao= 65"
  '
  '                  On Error Resume Next
  '                  Err.Clear
  '                  db.Execute SQL, afetados
  '                  If Err <> 0 Then
  '                      Call GeraLogs(CaminhoLogAgenteStartFlow, "[ERR] - Erro ao processar Passo: " & tmpTipo & ". SQL:" & SQL, lstLOG, True)
  '                      db.RollbackTrans
  '                      GoTo ProximoRegistro
  '                  End If
   '
  '                  On Error GoTo ErrPassoEspecificacao
  '
  '                  On Error Resume Next
  '                  Err.Clear
  '
  '                  objTaskMessage.ReplyToTask
  '
  '                  If Err.Number <> 0 Then
  '                      Call GeraLogs(CaminhoLogAgenteStartFlow, "[MSG] - Passo não respondido. Erro no ReplyToTask. Passo: " & tmpTipo, lstLOG, True)
  '                      db.RollbackTrans
  '                      GoTo ProximoRegistro
  '                  End If
  '
'                    Err.Clear
'                    objMsgColl.Item(1).Delete
'
'                    If Err.Number <> 0 Then
 '                       Call GeraLogs(CaminhoLogAgenteStartFlow, "[MSG] - Passo não respondido. Erro no Item.Delete. Passo: " & tmpTipo, lstLOG, True)
'                        db.RollbackTrans
'                        GoTo ProximoRegistro
'                    End If
'
'                    'On Error GoTo ErrPassoEspecificacao
'
'                    Call GeraLogs(CaminhoLogAgenteStartFlow, "[MSG] - Passo: " & tmpTipo & ", respondido com sucesso na rotina [ PassoEspecificacao() ].", lstLOG, True)
'                Else
'                    Call GeraLogs(CaminhoLogAgenteStartFlow, "[MSG] - STATUS_RESPOSTA não foi definido para ser processado no workflow. Passo: " & tmpTipo & " Ação: " & Trim(rs("status_resposta")), lstLOG, True)
'                    db.RollbackTrans
'                    If Not SendMail(EnviarEmailPara & ";", "", "", "Mensagem enviada pelo Agente Emissão-SAC Yasuda.[ERR3]-Falha no método IPSCall da Sysmap (IPS). Contate o Adminstrador do Sistema." & Err & " - " & Error & " - PI: " & rs!Num_PI) Then
 '                       Call GeraLogs(CaminhoLogAgenteStartFlow, "[ERR]- Erro ao enviar e-mail de aviso do agente de emissão: " & EnviarEmailPara, lstLOG, True)
'                    End If
'                    GoTo ProximoRegistro
'                End If
 '           End If
 '       Else
'            Call LiberaPITerminal(intRet, rs("num_pi"), lnum_prop, rs("cod_ramo"))
'            If intRet <> 0 Then
 '
 '               Call GeraLogs(CaminhoLogAgenteStartFlow, "Falha no método IPSCall da Sysmap (IPS)", lstLOG, True)
 '               db.RollbackTrans
 '               If Not SendMail(EnviarEmailPara & ";", "", "", "Mensagem enviada pelo Agente Emissão-SAC Yasuda.[ERR3]-Falha no método IPSCall da Sysmap (IPS). Contate o Adminstrador do Sistema." & Err & " - " & Error & " - PI: " & rs!Num_PI) Then
 '                   Call GeraLogs(CaminhoLogAgenteStartFlow, "[ERR]- Erro ao enviar e-mail de aviso do agente de emissão: " & EnviarEmailPara, lstLOG, True)
 '               End If
 '               GoTo ProximoRegistro
'            End If
 '       End If
 '
'        rs2.Close
 '       Set rs2 = Nothing
 '
'        SQL = "update tab_ctrl_espec "
'        SQL = SQL & " set dat_anexo  = " & Format(Date, "yyyymmdd")
 '       SQL = SQL & ", nomearq  = '" & NomeArq & "'"
 '       SQL = SQL & " Where Num_PI = " & Trim(rs("num_pi"))
'        db.Execute SQL
 '
'        If Err <> 0 Then
 '           Call GeraLogs(CaminhoLogAgenteStartFlow, "[ERR] - Falha ao atualizar registro na tabela tab_ctrl_espec. Descrição: " & Err.Description, lstLOG, True)
'            db.RollbackTrans
'            GoTo ProximoRegistro
'        End If
'
'        SQL = ""
'        SQL = "update tab_prop_mensagem "
'        SQL = SQL & " set dat_liber  = " & Format(Date, "yyyymmdd")
'        SQL = SQL & ", hor_liber  = " & Format(Time, "hhmmss")
'        SQL = SQL & " Where Num_PI = " & Trim(rs("num_pi"))
'        SQL = SQL & " and dat_liber = 0 "
'        SQL = SQL & " and cod_mensagem = 9559"
 '
'        db.Execute SQL
'        If Err <> 0 Then
'            Call GeraLogs(CaminhoLogAgenteStartFlow, "[ERR] - Falha ao atualizar registro na tabela TAB_PROP_MENSAGEM. Descrição: " & Err.Description, lstLOG, True)
'            db.RollbackTrans
'            GoTo ProximoRegistro
'        End If
'
'        db.CommitTrans
'
'ProximoRegistro:
'        On Error GoTo ErrPassoEspecificacao: Err.Clear
'        rs.MoveNext
 '       hasTrasaction = False
'    Loop
'
'    rs.Close
'    Set rs = Nothing
'
    Set objFolder = Nothing
    Set objFolder = Nothing
    Set objMsgColl = Nothing
    Set objMsgFilter = Nothing
    Set objTaskMessage = Nothing
'
'    PassoEspecificacao = True
'    Call GeraLogs(CaminhoLogAgenteStartFlow, "[MSG] - Fim da Rotina [ PassoEspecificacao() ] ...", lstLOG, True)
'
'
'
End Sub


