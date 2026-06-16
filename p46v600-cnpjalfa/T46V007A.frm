VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT46V007A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Localização dos dados"
   ClientHeight    =   5625
   ClientLeft      =   675
   ClientTop       =   1110
   ClientWidth     =   9015
   ControlBox      =   0   'False
   Icon            =   "T46V007A.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   Begin Threed.SSFrame fraX 
      Height          =   1200
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9000
      _Version        =   65536
      _ExtentX        =   15875
      _ExtentY        =   2117
      _StockProps     =   14
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Begin VB.Label lblX 
         Alignment       =   2  'Center
         Caption         =   "Estes dados podem ser compartilhados por todos os usuários, desde que estejam em ambiente de rede."
         ForeColor       =   &H00800000&
         Height          =   300
         Index           =   1
         Left            =   100
         TabIndex        =   3
         Top             =   850
         Width           =   8800
      End
      Begin VB.Label lblX 
         Alignment       =   2  'Center
         Caption         =   "O sistema instalado neste microcomputador está utilizando dados de cotações e propostas localizados em"
         ForeColor       =   &H00800000&
         Height          =   300
         Index           =   0
         Left            =   100
         TabIndex        =   1
         Top             =   200
         Width           =   8800
      End
      Begin VB.Label lblOrigem 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Local atual dos dados"
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   100
         TabIndex        =   2
         Top             =   500
         Width           =   8800
      End
   End
   Begin Threed.SSPanel panAviso 
      Height          =   500
      Left            =   100
      TabIndex        =   7
      Top             =   3500
      Visible         =   0   'False
      Width           =   8800
      _Version        =   65536
      _ExtentX        =   15522
      _ExtentY        =   882
      _StockProps     =   15
      ForeColor       =   8388608
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
      FloodColor      =   8454143
   End
   Begin Threed.SSFrame fraX 
      Height          =   2100
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   1200
      Width           =   9000
      _Version        =   65536
      _ExtentX        =   15875
      _ExtentY        =   3704
      _StockProps     =   14
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSFrame fraDestino 
         Height          =   500
         Index           =   1
         Left            =   2100
         TabIndex        =   14
         Top             =   1300
         Width           =   6800
         _Version        =   65536
         _ExtentX        =   11994
         _ExtentY        =   882
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox txtDestino 
            Height          =   285
            Left            =   100
            TabIndex        =   17
            Top             =   150
            Width           =   6300
         End
         Begin Threed.SSCommand cmdDestino 
            Height          =   240
            Index           =   1
            Left            =   6400
            TabIndex        =   18
            TabStop         =   0   'False
            ToolTipText     =   "Busca protocolo."
            Top             =   170
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   423
            _StockProps     =   78
            ForeColor       =   12582912
         End
      End
      Begin Threed.SSFrame fraDestino 
         Height          =   500
         Index           =   0
         Left            =   2100
         TabIndex        =   13
         Top             =   750
         Width           =   6800
         _Version        =   65536
         _ExtentX        =   11994
         _ExtentY        =   882
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin Threed.SSCommand cmdDestino 
            Height          =   240
            Index           =   0
            Left            =   6400
            TabIndex        =   15
            TabStop         =   0   'False
            ToolTipText     =   "Busca protocolo."
            Top             =   170
            Width           =   240
            _Version        =   65536
            _ExtentX        =   423
            _ExtentY        =   423
            _StockProps     =   78
            ForeColor       =   12582912
         End
         Begin VB.Label lblDestino 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Novo local para os dados"
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   100
            TabIndex        =   16
            Top             =   150
            Width           =   6300
         End
      End
      Begin VB.OptionButton optPasta 
         Caption         =   "Não alterar localização dos dados"
         ForeColor       =   &H00800000&
         Height          =   500
         Index           =   0
         Left            =   100
         TabIndex        =   12
         Top             =   200
         Width           =   3000
      End
      Begin VB.OptionButton optPasta 
         Caption         =   "Transferir dados para uma nova pasta"
         ForeColor       =   &H00800000&
         Height          =   500
         Index           =   2
         Left            =   100
         TabIndex        =   6
         Top             =   1300
         Width           =   2000
      End
      Begin VB.OptionButton optPasta 
         Caption         =   "Transferir dados para uma pasta existente"
         ForeColor       =   &H00800000&
         Height          =   500
         Index           =   1
         Left            =   100
         TabIndex        =   5
         Top             =   750
         Width           =   2000
      End
   End
   Begin Threed.SSCommand cmdOk 
      Height          =   300
      Left            =   4000
      TabIndex        =   11
      Top             =   5200
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1764
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Ok"
      ForeColor       =   8388608
   End
   Begin Threed.SSFrame fraX 
      Height          =   900
      Index           =   2
      Left            =   0
      TabIndex        =   8
      Top             =   4100
      Width           =   9000
      _Version        =   65536
      _ExtentX        =   15875
      _ExtentY        =   1587
      _StockProps     =   14
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
      Alignment       =   2
      Begin VB.Label lblX 
         Alignment       =   2  'Center
         Caption         =   "utilizando a opção de menu Ferramentas + Localização dos dados."
         ForeColor       =   &H00800000&
         Height          =   300
         Index           =   2
         Left            =   100
         TabIndex        =   10
         Top             =   500
         Width           =   8800
      End
      Begin VB.Label lblX 
         Alignment       =   2  'Center
         Caption         =   "A qualquer momento este processo de localização dos dados pode ser acionado,"
         ForeColor       =   &H00800000&
         Height          =   300
         Index           =   3
         Left            =   100
         TabIndex        =   9
         Top             =   200
         Width           =   8800
      End
   End
End
Attribute VB_Name = "frmT46V007A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mErrCarga As Boolean    'Indica que houve erro na carga.
Private mNomBasDados() As String  'Nomes das bases de dados.
Private mP00Tarifa As Byte    'Indica se a base de dados P00Tarifa está no mesmo endereço das demais
'bases de dados: 1 = Sim e 2 = Não.
Private mQuaBasDados As Byte  'Quantidade de bases de dados.
Private Function mfCopBasDados(ByVal pOrigem As String, ByVal pDestino As String)
'Função: copia arquivos do local de origem para o local de destino.

'Parâmetros de entrada...pOrigem....Local de origem do arquivo.
'                        pDestino...Local de destino do arquivo.

    Dim lEndDestino As String   'Endereço de destino.
    Dim lEndOrigem As String  'Endereço de origem.
    Dim lI As Byte    'Índice.
    Dim lMensagem As String    'Mensagem.
    Dim lOpcao As Byte    'Opção do form T28V011A.

    mfCopBasDados = False

    '1. Posiciona controles.
    If Right$(pOrigem, 1) <> "\" Then
        pOrigem = pOrigem & "\"
    End If
    If Right$(pDestino, 1) <> "\" Then
        pDestino = pDestino & "\"
    End If

    '2. Copia arquivos para destino.
    panAviso.Visible = True
    panAviso.Refresh
    For lI = 1 To mQuaBasDados
        'a) Posiciona endereços de origem e de destino.
        lEndOrigem = pOrigem & mNomBasDados(lI) & ".MDB"
        lEndDestino = pDestino & mNomBasDados(lI) & ".MDB"

        'b) Para o primeiro processamento, verifica se já existe o arquivo de destino.
        If lI = 1 Then
            If gM46V999.gfPreenchido(Dir(lEndDestino)) Then
                With frmT46V008A
                    .lblOrigem.Caption = pOrigem
                    .lblDestino.Caption = pDestino
                    .Show vbModal, Me
                    If .optLocalizacao(0) = True Then
                        lOpcao = 1
                    Else
                        If .optLocalizacao(1) = True Then
                            lOpcao = 2
                        Else
                            If .optLocalizacao(2) = True Then
                                lOpcao = 3
                            Else
                                lOpcao = 0
                            End If
                        End If
                    End If
                End With
                Unload frmT46V008A
                Select Case lOpcao
                Case 0
                    'Opção não selecionada.
                    Call gM46V999.gpGraLog(1, "FNP0029 - Opção não selecionada.")
                    Exit Function
                Case 1
                    'Excluir os dados do local de destino e substituir pelos dados do local de origem.
                    If mfRenBasDados(pDestino) = False Then
                        Exit Function
                    End If
                Case 2
                    'Manter os dados existentes no local de destino e excluir os dados do local de
                    'origem.
                    Exit For
                Case 3
                    'Interromper o processo.
                    Call gM46V999.gpGraLog(0, lMensagem & " CANCELAR")
                    Exit Function
                End Select
            End If
        End If

        'c) Posiciona quadro de avisos.
        panAviso.Caption = "Copiando de: " & lEndOrigem & "." & vbLf & "para: " & lEndDestino & "."
        panAviso.Refresh

        'd) Copia base de dados da origem para o destino
        If mNomBasDados(lI) = "P00Tarifa" And mP00Tarifa = 2 Then
            'Se o arquivo P0044800 não estiver na mesma pasta dos demais arquivos, não copiar.
        Else
            On Error Resume Next
            FileCopy lEndOrigem, lEndDestino
            If Err <> 0 Then
                Call gM46V999.gpGraLog(2, "FTE0070 - Erro ao transferir base de dados " & lEndOrigem & " para " & _
                                          lEndDestino & " (processo de localização dos dados: " & Err & " - " & _
                                          Error & ").")
                On Error GoTo 0
                Exit Function
            End If
            On Error GoTo 0
        End If
    Next lI

    '3. Altera nome das bases de dados no diretório de origem.
    If mfRenBasDados(pOrigem) = False Then
        Exit Function
    End If

    mfCopBasDados = True
End Function
Private Function mfRenBasDados(ByVal pLocal As String) As Boolean
'Função: altera nome das bases de dados com extensão .MDB para extensão .OLD.

    Dim lEnderecoMDB As String    'Endereço original da base de dados.
    Dim lEnderecoOLD As String    'Endereço da base de dados com extensão OLD.
    Dim lI As Byte             'Usado com For...Next.

    mfRenBasDados = False

    For lI = 1 To mQuaBasDados
        If mNomBasDados(lI) = "P00Tarifa" And mP00Tarifa = 2 Then
            'Não renomear o arquivo.
        Else
            '1. Posiciona nomes.
            If Right$(pLocal, 1) = "\" Then
                lEnderecoMDB = pLocal & mNomBasDados(lI) & ".MDB"
                lEnderecoOLD = pLocal & mNomBasDados(lI) & ".OLD"
            Else
                lEnderecoMDB = pLocal & "\" & mNomBasDados(lI) & ".MDB"
                lEnderecoOLD = pLocal & "\" & mNomBasDados(lI) & ".OLD"
            End If

            '2. Renomeia arquivo.
            If Not gM46V999.gfPreenchido(Dir(lEnderecoMDB)) Then
                Call gM46V999.gpGraLog(2, "FTE0054 - Erro ao alterar nome do arquivo " & lEnderecoMDB & _
                                        " para " & lEnderecoOLD & ": base de dados " & lEnderecoMDB & _
                                        " não localizada.")
                Exit Function
            End If
            On Error Resume Next
            Name lEnderecoMDB As lEnderecoOLD
            If Err <> 0 Then
                On Error GoTo 0
                On Error Resume Next
                Kill lEnderecoOLD
                On Error GoTo 0
                On Error Resume Next
                Name lEnderecoMDB As lEnderecoOLD
                If Err <> 0 Then
                    Call gM46V999.gpGraLog(2, "FTE0054 - Erro ao alterar nome do arquivo " & lEnderecoMDB & _
                                            " para " & lEnderecoOLD & ": " & Err & " - " & Error)
                    On Error GoTo 0
                    Exit Function
                End If
                On Error GoTo 0
            End If
        End If
    Next lI

    mfRenBasDados = True
End Function
Private Sub cmdDestino_Click(Index As Integer)
    Dim clsA99V002 As clsA99V002
    Dim lMensagem As String    'Mensagem.

    '1. Seleciona destino.
    Set clsA99V002 = New clsA99V002
    With clsA99V002
        Select Case Index
        Case 1      'Pasta existente.
            lMensagem = "Selecionar pasta existente para transferir as bases de dados"
            lblDestino.Caption = clsA99V002.gfObtDiretorio(lMensagem) & "\"
        Case 2      'Nova pasta.
            txtDestino.Text = clsA99V002.gfObtDiretorio("Selecionar endereço onde será criada " & _
                                                        "pasta para transferir as bases de dados")
        End Select
    End With
    Set clsA99V002 = Nothing
End Sub
Private Sub cmdOK_Click()
    Dim lAltLinha As Boolean    'Indicador: alterar linha?
    Dim lAppPath As Boolean    'Indica que o endereço de origem é o APP.Path.
    Dim lDestino As String    'Local de destino dos arquivos.
    Dim lEndSyasMultiINI As String   'Endereço do arquivo SyasMulti.INI.
    Dim lEndSyasMultiNEW As String   'Endereço do arquivo SyasMulti.NEW.
    Dim lEndSyasMultiTMP As String   'Endereço do arquivo SyasMulti.TMP.
    Dim lI As Byte             'Índice.
    Dim lLinha As String    'Linha de registro.
    Dim lMensagem As String    'Mensagem.
    Dim lNumSyasMultiINI As Integer  'Utilizado com FreeFile para SyasMulti.INI.
    Dim lNumSyasMultiNEW As Integer  'Utilizado com FreeFile para SyasMulti.NEW.
    Dim lOpcao As Byte    'Opção.
    Dim lOrigem As String    'Local de origem dos arquivos.
    Dim lPosicao As Integer    'Utilizado com InStr.
    Dim lTamanho As Integer    'Tamanho.
    Dim lTipo As String    'Tipo.
    Dim lTitulo As String    'Título.

    Screen.MousePointer = vbHourglass

    '1. Verifica opção de pasta.
    If optPasta(0) = True Then      'Não alterar localização.
        Unload Me
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If optPasta(1) = True Then
        lOpcao = 1      'Pasta existente.
    Else
        If optPasta(2) = True Then
            lOpcao = 2  'Criar pasta.
        Else
            Call gM46V999.gpGraLog(1, "FNP0029 - Opção não selecionada.")
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If

    '2. Verifica preenchimento do endereço do novo local.
    lOrigem = Trim$(lblOrigem.Caption)
    Select Case lOpcao
    Case 1      'Pasta existente.
        If Not gM46V999.gfPreenchido(lblDestino.Caption) Then
            Call gM46V999.gpGraLog(1, "FNP0030 - Endereço de destino para transferência dos dados não " & _
                                      "selecionado.")
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        lDestino = Trim$(lblDestino.Caption)
    Case 2      'Nova pasta.
        If Not gM46V999.gfPreenchido(txtDestino.Text) Then
            Call gM46V999.gpGraLog(1, "FNP0031 - Endereço de destino para transferência dos dados não " & _
                                      "preenchido.")
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        lDestino = Trim$(txtDestino.Text)
    End Select
    '2. Posiciona controle dos endereços dos arquivos (gM46V999.gAppPath).
    If Right$(lOrigem, 1) <> "\" Then
        lOrigem = lOrigem & "\"
    End If
    If Right$(lDestino, 1) <> "\" Then
        lDestino = lDestino & "\"
    End If
    If UCase$(lOrigem) = UCase$(lDestino) Then
        Call gM46V999.gpGraLog(1, "FGE0058 - Endereço de origem dos dados igual ao endereço de destino: " & _
                                  lOrigem & ".")
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    '3. Obtém confirmação do usuário.
    If lOrigem = App.Path & "\" Then
        lAppPath = True
    Else
        lAppPath = False
    End If
    While Right$(lDestino, 1) = "\" Or Right$(lDestino, 1) = ":"
        lDestino = Left$(lDestino, Len(lDestino) - 1)
    Wend
    If Len(lDestino) <= 2 Then
        Call gM46V999.gpGraLog(1, "FGE0068 - Endereço de destino inválido (não pode ser a raiz de um " & _
                                  "diretório): " & lDestino & ".")
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    lMensagem = "AGE0043 - Confirma a alteração da localização dos dados?" & vbLf & vbLf & _
                "DE: " & lOrigem & vbLf & vbLf & "PARA: " & lDestino
    If MsgBox(lMensagem, vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
        Call gM46V999.gpGraLog(0, lMensagem & " SIM")
    Else
        Call gM46V999.gpGraLog(0, lMensagem & " NÃO")
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    '4. Verifica existência do endereço de destino.
    If gM46V999.gfPreenchido(Dir(lDestino, vbDirectory)) Then
        If lOpcao = 2 Then
            Call gM46V999.gpGraLog(1, "FGE0067 - Endereço de destino já existente: " & lDestino & "." & vbLf & _
                                      vbLf & "Para transferir para este endereço, selecionar opção " & _
                                      """Transferir dados para uma pasta existente"".")
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    Else
        If lOpcao = 1 Then
            Call gM46V999.gpGraLog(1, "FGE0065 - Endereço de destino não localizado: " & lDestino & "." & vbLf & _
                                      vbLf & "Para transferir para este endereço, selecionar opção " & _
                                      """Transferir dados para uma nova pasta"".")
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If

    '5. Cria pasta.
    If lOpcao = 2 Then      'Cria pasta.
        On Error Resume Next
        MkDir lDestino
        If Err <> 0 Then
            On Error GoTo 0
            Call gM46V999.gpGraLog(2, "FTE0033 - Erro ao criar pasta " & lDestino & _
                                    " (pasta para transferência dos dados - " & Err & " - " & Error & ").")
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        On Error GoTo 0
    End If

    '6. Copia arquivos do local atual para o novo local.
    If mfCopBasDados(lOrigem, lDestino) = False Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    '7. Posiciona nomes dos arquivos de inicialização.
    lEndSyasMultiINI = gM46V999.gAppPath & "SyasMult.INI"
    lEndSyasMultiNEW = gM46V999.gAppPath & "SyasMult.NEW"
    lEndSyasMultiTMP = gM46V999.gAppPath & "SyasMult.TMP"

    '8. Abre arquivo de inicialização.
    If Not gM46V999.gfPreenchido(Dir(lEndSyasMultiINI)) Then
        Call gM46V999.gpGraLog(2, "FTE0025 - Arquivo não localizado: " & lEndSyasMultiINI & _
                                " (processo de localização de dados).")
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    lNumSyasMultiINI = FreeFile
    On Error Resume Next
    Open lEndSyasMultiINI For Input Access Read As lNumSyasMultiINI
    If Err <> 0 Then
        Call gM46V999.gpGraLog(2, "FTE0026 - Erro ao abrir arquivo " & lEndSyasMultiINI & ": " & Err & " - " & _
                                  Error)
        On Error GoTo 0
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    On Error GoTo 0

    '9. Altera arquivo de inicialização.
    lDestino = lDestino & "\"
    lNumSyasMultiNEW = FreeFile
    On Error Resume Next
    Open lEndSyasMultiNEW For Output As #lNumSyasMultiNEW
    lAltLinha = False
    Do While Not EOF(lNumSyasMultiINI)
        Line Input #lNumSyasMultiINI, lLinha
        If gM46V999.gfPreenchido(lLinha) Then
            lLinha = Trim$(lLinha)
            Select Case UCase$(Left$(lLinha, 1))
            Case ";"    'Linha de observação.
                'Nada a processar - linha de observação.
            Case "["    'Linha de título.
                lTitulo = UCase$(lLinha)
            Case Else   'Linha de tipo.
                Select Case lTitulo
                Case "[P00APOL]", "[P00MPED]"
                    lPosicao = InStr(lLinha, "=")
                    lTamanho = Len(lLinha)
                    If lPosicao = 0 Then
                        Call gM46V999.gpGraLog(2, "FTE0009 - Registro inválido no arquivo " & _
                                                  lEndSyasMultiINI & _
                                                " (sinal de igual não localizado)." & vbLf & vbLf & _
                                                  lLinha)
                        Exit Sub
                    End If
                    If lPosicao = 1 Then
                        Call gM46V999.gpGraLog(2, "FTE0009 - Registro inválido no arquivo " & _
                                                  lEndSyasMultiINI & _
                                                " (sinal de igual na primeira posição)." & vbLf & _
                                                  vbLf & lLinha)
                        Exit Sub
                    End If
                    If lPosicao = lTamanho Then
                        Call gM46V999.gpGraLog(2, "FTE0009 - Registro inválido no arquivo " & _
                                                  lEndSyasMultiINI & _
                                                " (sinal de igual na última posição)." & vbLf & _
                                                  vbLf & lLinha)
                        Exit Sub
                    End If
                    lTipo = Trim$(UCase$(Left$(lLinha, lPosicao - 1)))
                    Select Case lTipo
                    Case "TIPO", "CARGA"
                        'Tipos existentes, mas não necessário neste processo.
                    Case "DATABASE"
                        lAltLinha = True
                    Case Else
                        Call gM46V999.gpGraLog(2, "FTE0009 - Registro inválido no arquivo " & _
                                                  lEndSyasMultiINI & _
                                                " (tipo desconhecido para título " & lTitulo & _
                                                  ")." & vbLf & vbLf & lLinha)
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End Select
                End Select
                If lAltLinha = True Then
                    lPosicao = InStr(lLinha, lOrigem)
                    If lPosicao = 0 Then
                        If lAppPath = True Then
                            lLinha = Left$(lLinha, 9) & lDestino & Mid$(lLinha, 10)
                        Else
                            Call gM46V999.gpGraLog(2, "FTE0009 - Registro inválido no arquivo " & _
                                                      lEndSyasMultiINI & " (local diferente de " & lOrigem & _
                                                      ")." & vbLf & vbLf & lLinha)
                            Screen.MousePointer = vbDefault
                            Exit Sub
                        End If
                    Else
                        lLinha = Left$(lLinha, lPosicao - 1) & lDestino & _
                                 Mid$(lLinha, lPosicao + Len(lOrigem))
                    End If
                    lAltLinha = False
                End If
            End Select
        End If
        Print #lNumSyasMultiNEW, lLinha
    Loop
    Close lNumSyasMultiINI
    Close lNumSyasMultiNEW

    '10. Exclui arquivo INI e renomeia arquivo NEW para INI.
    Name lEndSyasMultiINI As lEndSyasMultiTMP
    Name lEndSyasMultiNEW As lEndSyasMultiINI
    Kill lEndSyasMultiTMP

    '11. Reprocessa arquivo de inicialização.
    Call gM46V999.gpInicia("")
    frmT46V001A.panBasDados.Caption = lDestino

    '12. Encerra o processo.
    Screen.MousePointer = vbDefault
    Unload Me
End Sub
Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    frmT46V001A.panForm = Mid$(Me.Name, 4)
    If mErrCarga = True Then
        Unload Me
    End If
End Sub
Private Sub Form_Load()
    Dim lEndBasDados() As String  'Endereço das bases de dados.
    Dim lEndereco As String    'Endereço.
    Dim lEndP00MPED As String    'Endereço da base de dados P00MPED.MDB.
    Dim lI As Byte             'Índice.
    Dim lPosicao As Integer    'Posição - utilizado com InStr.

    Screen.MousePointer = vbHourglass

    '1. Posiciona e centraliza form.
    panAviso.Visible = False
    Call gM46V999.gpCenForm(Me)

    '2. Posiciona controles.
    mErrCarga = False
    mQuaBasDados = 2       '<<<<<<====== Posicionar a quantidade de bases de dados.
    ReDim mNomBasDados(mQuaBasDados)
    ReDim lEndBasDados(mQuaBasDados)
    mNomBasDados(1) = "P00MPED"
    lEndBasDados(1) = gM46V999.gEndP00MPED
    mNomBasDados(2) = "P00Apol"
    lEndBasDados(2) = gM46V999.gEndP00APOL


    '3. Verifica se as bases de dados estão na mesma pasta (a base de dados P0044800 pode estar numa
    '   pasta diferente; neste caso, posicionar a variável mP0044800).
    mP00Tarifa = 1    'Indica que a base de dados P0044800 está na mesma pasta das demais bases de dados.
    For lI = 1 To mQuaBasDados
        lPosicao = InStr(UCase(lEndBasDados(lI)), UCase(mNomBasDados(lI)))
        If lI = 1 Then
            '===1) Primeira base de dados (P0042100): posiciona endereço para verificação.
            If lPosicao = 1 Then
                lEndP00MPED = gM46V999.gAppPath
            Else
                lEndP00MPED = Left$(gM46V999.gEndP00MPED, lPosicao - 1)
            End If
        Else
            '===1) Posiciona endereço da base de dados.
            If lPosicao = 1 Then
                lEndereco = gM46V999.gAppPath
            Else
                lEndereco = Left$(lEndBasDados(lI), lPosicao - 1)
            End If
            '===2) Verifica se está na mesma pasta da base de dados P0042100.
            'If UCase$(lEndP00MPED) <> UCase$(lEndereco) Then
            '    If mNomBasDados(lI) = "P00Tarifa" Then
            '        mP00Tarifa = 2
            '    Else
            '        Call gM46V999.gpGraLog(2, "FTE0069 - Localização da base de dados " & mNomBasDados(lI) & _
                     '                         ".MDB (" & lEndBasDados(lI) & _
                     '                         ") diferente da localização da base de dados P00MPED.MDB (" & _
                     '                         gM46V999.gEndP00MPED & ")." & vbLf & vbLf & "Processo interrompido.")
            '        mErrCarga = True
            '        Exit Sub
            '    End If
            'End If
        End If
    Next lI

    '4. Posiciona campos do form.
    lblOrigem.Caption = lEndP00MPED
    optPasta(0) = True      'Não alterar localização dos dados.
    lblDestino.Caption = ""
    fraDestino(0).Visible = False
    txtDestino.Text = ""
    fraDestino(1).Visible = False
End Sub

Private Sub optPasta_Click(Index As Integer)
    Select Case Index
    Case 0  'Não alterar localização dos dados.
        fraDestino(0).Visible = False
        fraDestino(1).Visible = False
    Case 1  'Pasta existente.
        fraDestino(0).Visible = True
        fraDestino(1).Visible = False
        Call cmdDestino_Click(Index)
    Case 2  'Criar pasta.
        fraDestino(0).Visible = False
        fraDestino(1).Visible = True
        Call cmdDestino_Click(Index)
    End Select
End Sub
