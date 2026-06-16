VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT46V190A 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dados de contato para Inspeção"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   8115
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame 
      Caption         =   "Contatos"
      ForeColor       =   &H00800000&
      Height          =   1365
      Left            =   60
      TabIndex        =   16
      Top             =   1590
      Width           =   7935
      Begin VB.TextBox txtTelefone 
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   6555
         MaxLength       =   14
         TabIndex        =   20
         ToolTipText     =   "Telefone da inspeção. Formato: (99)9999-9999"
         Top             =   270
         Width           =   1240
      End
      Begin VB.TextBox txtContato 
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   885
         MaxLength       =   40
         TabIndex        =   18
         ToolTipText     =   "Contato de inspeção."
         Top             =   270
         Width           =   4620
      End
      Begin VB.TextBox txtTelefone 
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   1
         Left            =   6555
         MaxLength       =   14
         TabIndex        =   24
         ToolTipText     =   "Telefone da inspeção. Formato: (99)9999-9999"
         Top             =   600
         Width           =   1240
      End
      Begin VB.TextBox txtContato 
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   1
         Left            =   885
         MaxLength       =   40
         TabIndex        =   22
         ToolTipText     =   "Contato de inspeção."
         Top             =   600
         Width           =   4620
      End
      Begin VB.TextBox txtTelefone 
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   2
         Left            =   6555
         MaxLength       =   14
         TabIndex        =   28
         ToolTipText     =   "Telefone da inspeção. Formato: (99)9999-9999"
         Top             =   930
         Width           =   1240
      End
      Begin VB.TextBox txtContato 
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   2
         Left            =   885
         MaxLength       =   40
         TabIndex        =   26
         ToolTipText     =   "Contato de inspeção."
         Top             =   930
         Width           =   4620
      End
      Begin VB.Label lblTelefone 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Telefone 1: "
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5580
         TabIndex        =   19
         Top             =   270
         Width           =   975
      End
      Begin VB.Label lblContato 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nome 1:"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   150
         TabIndex        =   17
         Top             =   270
         Width           =   735
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Telefone 2: "
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5610
         TabIndex        =   23
         Top             =   600
         Width           =   945
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nome 2:"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   150
         TabIndex        =   21
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Telefone 3: "
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5610
         TabIndex        =   27
         Top             =   930
         Width           =   945
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nome 3:"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   150
         TabIndex        =   25
         Top             =   930
         Width           =   735
      End
   End
   Begin VB.TextBox txtEmail 
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1050
      MaxLength       =   60
      TabIndex        =   30
      ToolTipText     =   "Contato de inspeção."
      Top             =   3060
      Width           =   6960
   End
   Begin VB.TextBox txtPontoReferencia 
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   1050
      MaxLength       =   200
      MultiLine       =   -1  'True
      TabIndex        =   32
      Top             =   3420
      Width           =   6960
   End
   Begin Threed.SSCommand cmdSalvar 
      Height          =   480
      Left            =   60
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   4110
      Width           =   1260
      _Version        =   65536
      _ExtentX        =   2222
      _ExtentY        =   847
      _StockProps     =   78
      Caption         =   "Salvar"
      ForeColor       =   12582912
      Picture         =   "T46V190A.frx":0000
   End
   Begin Threed.SSCommand cmdCopiarDadoInspecao 
      Height          =   480
      Left            =   1600
      TabIndex        =   34
      TabStop         =   0   'False
      ToolTipText     =   "Copiar dados de inspeção para os demais itens."
      Top             =   4110
      Width           =   1260
      _Version        =   65536
      _ExtentX        =   2222
      _ExtentY        =   847
      _StockProps     =   78
      Caption         =   "Replicar"
      ForeColor       =   12582912
      Picture         =   "T46V190A.frx":001C
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Email:"
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   60
      TabIndex        =   29
      Top             =   3060
      Width           =   975
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ponto de Referência:"
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   60
      TabIndex        =   31
      Top             =   3420
      Width           =   975
   End
   Begin VB.Label lblCEP 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Segurado"
      Height          =   300
      Left            =   6660
      TabIndex        =   11
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CEP: "
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   6120
      TabIndex        =   10
      Top             =   840
      Width           =   525
   End
   Begin VB.Label lblCidade 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Segurado"
      Height          =   300
      Left            =   3300
      TabIndex        =   15
      Top             =   1200
      Width           =   4695
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cidade: "
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   2430
      TabIndex        =   14
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblUF 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Segurado"
      Height          =   300
      Left            =   1020
      TabIndex        =   13
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "UF: "
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   60
      TabIndex        =   12
      Top             =   1200
      Width           =   945
   End
   Begin VB.Label lblComplemento 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Segurado"
      Height          =   300
      Left            =   1020
      TabIndex        =   9
      Top             =   840
      Width           =   4575
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Complem.: "
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   60
      TabIndex        =   8
      Top             =   840
      Width           =   945
   End
   Begin VB.Label lblNumEndereco 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Segurado"
      Height          =   300
      Left            =   6660
      TabIndex        =   7
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nº: "
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   6210
      TabIndex        =   6
      Top             =   480
      Width           =   435
   End
   Begin VB.Label lblNomeSegurado 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Segurado"
      Height          =   300
      Left            =   3270
      TabIndex        =   3
      Top             =   120
      Width           =   4725
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Segurado: "
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   2400
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblEndereco 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Segurado"
      Height          =   300
      Left            =   1020
      TabIndex        =   5
      Top             =   480
      Width           =   5055
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Endereço: "
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   60
      TabIndex        =   4
      Top             =   480
      Width           =   945
   End
   Begin VB.Label lblNumItem 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Segurado"
      Height          =   300
      Left            =   1020
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Item: "
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   945
   End
End
Attribute VB_Name = "frmT46V190A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public TodosItensCorretos As Boolean
Public NossoNumeroOld As String

Private Tecla As Integer
Private PonteiroMouse As Long
Private TodosNossoNumero As String

''2016-04-12 : Juliano -  Replicar contatos de inspecao. Mantive a forma de atualizar/inserir
''             conforme feito anteriormente para manter a regra de validacao, reduzir tempo e
''             uniformidade(???) de cod. Tem espaco para melhorar o codigo.
''             Bug do telefone3 corrigido
Private Sub cmdCopiarDadoInspecao_Click()
    Dim lNome As String
    Dim lTelefone As String
    Dim lSql As String
    Dim bdBasDados As ADODB.Connection    'Base de dados.
    Dim lMensagem As String
    
    Screen.MousePointer = vbHourglass
    
    'Telefone 1
    lTelefone = Trim(txtTelefone(0).Text)
    If gM46V999.gfPreenchido(lTelefone) = False Then
        Call gM46V999.gpGraLog(1, "FGE0105 - Telefone 1 de contato para inspeção inválido (não preenchido).")
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        '1000477 - Telefone 1 de Contato Formatado
        '1 - valida se esta totalmente preenchido
        If Len(lTelefone) < 13 Then
            Call gM46V999.gpGraLog(1, "FGE0105 - Telefone 1 de contato para inspeção inválido (Favor informar o DDD e o número do Telefone 1 com 8 dígitos.)" & vbCrLf & "Formato: (99)9999-9999.")
            Screen.MousePointer = vbDefault
            Exit Sub
            
        Else
            '2 - valida formado (99)9999-9999
            If Mid(lTelefone, 1, 1) <> "(" Or Mid(lTelefone, 4, 1) <> ")" Or (Mid(lTelefone, 9, 1) <> "-" And Mid(lTelefone, 10, 1) <> "-") Then    '1200178 - Telefone 1 9 Digitos
                Call gM46V999.gpGraLog(1, "FGE0105 - Telefone 1 de contato para inspeção inválido (Favor informar o DDD e o número do Telefone 1 com 8 dígitos.)" & vbCrLf & "Formato: (99)9999-9999.")
                Screen.MousePointer = vbDefault
                Exit Sub
            Else
                '3 - não permite todos os numeros do Telefone 1 iguais
                If (InStr(1, "0000-0000 1111-1111 2222-2222 3333-3333 4444-4444 5555-5555 6666-6666 7777-7777 8888-8888 9999-9999", Mid(lTelefone, 5, Len(lTelefone)), vbTextCompare) > 0) Or _
                   (InStr(1, "00000-0000 11111-1111 22222-2222 33333-3333 44444-4444 55555-5555 66666-6666 77777-7777 88888-8888 99999-9999", Mid(lTelefone, 5, Len(lTelefone)), vbTextCompare) > 0) Then   ''1200178 - Telefone 1 9 Digitos
                    Call gM46V999.gpGraLog(1, "FGE0105 - Telefone 1 de contato para inspeção inválido (Favor informar o DDD e o número do Telefone 1 com 8 dígitos.)" & vbCrLf & "Formato: (99)99999-9999.")
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            End If
        End If
    End If
    
    'Telefone2
    lTelefone = Trim(txtTelefone(1).Text)
    If gM46V999.gfPreenchido(lTelefone) = False Then
        Call gM46V999.gpGraLog(1, "FGE0105 - Telefone 2 de contato para inspeção inválido (não preenchido).")
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        '1000477 - Telefone 2 de Contato Formatado
        '1 - valida se esta totalmente preenchido
        If Len(lTelefone) < 13 Then
            Call gM46V999.gpGraLog(1, "FGE0105 - Telefone 2 de contato para inspeção inválido (Favor informar o DDD e o número do Telefone 2 com 8 dígitos.)" & vbCrLf & "Formato: (99)9999-9999.")
            Screen.MousePointer = vbDefault
            Exit Sub
        Else
            '2 - valida formado (99)9999-9999
            If Mid(lTelefone, 1, 1) <> "(" Or Mid(lTelefone, 4, 1) <> ")" Or (Mid(lTelefone, 9, 1) <> "-" And Mid(lTelefone, 10, 1) <> "-") Then    '1200178 - Telefone 2 9 Digitos
                Call gM46V999.gpGraLog(1, "FGE0105 - Telefone 2 de contato para inspeção inválido (Favor informar o DDD e o número do Telefone 2 com 8 dígitos.)" & vbCrLf & "Formato: (99)9999-9999.")
                Screen.MousePointer = vbDefault
                Exit Sub
            Else
                '3 - não permite todos os numeros do Telefone 2 iguais
                If (InStr(1, "0000-0000 1111-1111 2222-2222 3333-3333 4444-4444 5555-5555 6666-6666 7777-7777 8888-8888 9999-9999", Mid(lTelefone, 5, Len(lTelefone)), vbTextCompare) > 0) Or _
                   (InStr(1, "00000-0000 11111-1111 22222-2222 33333-3333 44444-4444 55555-5555 66666-6666 77777-7777 88888-8888 99999-9999", Mid(lTelefone, 5, Len(lTelefone)), vbTextCompare) > 0) Then   ''1200178 - Telefone 2 9 Digitos
                    Call gM46V999.gpGraLog(1, "FGE0105 - Telefone 2 de contato para inspeção inválido (Favor informar o DDD e o número do Telefone 2 com 8 dígitos.)" & vbCrLf & "Formato: (99)99999-9999.")
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            End If
        End If
    End If
    
    'Telefone3 não é obrigatorio
    lTelefone = Trim(txtTelefone(2).Text)
    If gM46V999.gfPreenchido(lTelefone) Then
        '1000477 - Telefone 3 de Contato Formatado
        '1 - valida se esta totalmente preenchido
        If Len(lTelefone) < 13 Then
            Call gM46V999.gpGraLog(1, "FGE0105 - Telefone 3 de contato para inspeção inválido (Favor informar o DDD e o número do Telefone 3 com 8 dígitos.)" & vbCrLf & "Formato: (99)9999-9999.")
            Screen.MousePointer = vbDefault
            Exit Sub
        Else
            '2 - valida formado (99)9999-9999
            If Mid(lTelefone, 1, 1) <> "(" Or Mid(lTelefone, 4, 1) <> ")" Or (Mid(lTelefone, 9, 1) <> "-" And Mid(lTelefone, 10, 1) <> "-") Then    '1200178 - Telefone 3 9 Digitos
                Call gM46V999.gpGraLog(1, "FGE0105 - Telefone 3 de contato para inspeção inválido (Favor informar o DDD e o número do Telefone 3 com 8 dígitos.)" & vbCrLf & "Formato: (99)9999-9999.")
                Screen.MousePointer = vbDefault
                Exit Sub
            Else
                '3 - não permite todos os numeros do Telefone 3 iguais
                If (InStr(1, "0000-0000 1111-1111 2222-2222 3333-3333 4444-4444 5555-5555 6666-6666 7777-7777 8888-8888 9999-9999", Mid(lTelefone, 5, Len(lTelefone)), vbTextCompare) > 0) Or _
                   (InStr(1, "00000-0000 11111-1111 22222-2222 33333-3333 44444-4444 55555-5555 66666-6666 77777-7777 88888-8888 99999-9999", Mid(lTelefone, 5, Len(lTelefone)), vbTextCompare) > 0) Then   ''1200178 - Telefone 3 9 Digitos
                    Call gM46V999.gpGraLog(1, "FGE0105 - Telefone 3 de contato para inspeção inválido (Favor informar o DDD e o número do Telefone 3 com 8 dígitos.)" & vbCrLf & "Formato: (99)99999-9999.")
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            End If
        End If
    End If

    If (gM46V999.gfPreenchido(txtContato(0).Text) = False Or Len(txtContato(0).Text) <= 3) Or _
       (gM46V999.gfPreenchido(txtContato(1).Text) = False Or Len(txtContato(1).Text) <= 3) Then
        Call gM46V999.gpGraLog(1, "FGE0106 - Necessario informar ao menos dois contatos para inspeção válidos (informar no mínimo 3 caracteres).")
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    ' Salvar os dados
    If gM46V999.gfAbrBasDados("P00MPED", gM46V999.gTipP00MPED, gM46V999.gEndP00MPED, bdBasDados, lMensagem) = False Then
        Call gM46V999.gpGraLog(1, "Erro ao abrir base de dados - " & lMensagem)
        Unload Me
    End If
    
    ' Obter os dados digitados
    lNome = gM46V999.gfRetirarCaracterEspecial(Trim(txtContato(0).Text), False, True, True)
    lTelefone = Trim(txtTelefone(0).Text)
    
    lSql = "UPDATE      TPI" & vbCrLf
    lSql = lSql & "SET         PESS_CONT = " & gM46V999.gfFormatarCamposSQL(lNome) & "," & vbCrLf
    lSql = lSql & "            NUM_TEL   = " & gM46V999.gfFormatarCamposSQL(lTelefone) & "" & vbCrLf
    lSql = lSql & "FROM        RamosDiversos.dbo.Tab_Ped_Insp  TPI" & vbCrLf
    lSql = lSql & "INNER JOIN  RamosDiversos.dbo.Tab_Ped_Loc   TPL with (NOLOCK)" & vbCrLf
    lSql = lSql & "    ON      TPL.NOSSO_NUMERO                = TPI.NOSSO_NUMERO" & vbCrLf
    lSql = lSql & "        AND TPL.NUM_ITEM                    = TPI.NUM_ITEM" & vbCrLf
    lSql = lSql & "WHERE       TPL.NOSSO_NUMERO                IN (" & TodosNossoNumero & ")" & vbCrLf
    
    If gM46V999.gfExeSQL(bdBasDados, lSql, lMensagem) <> 0 Then
        Call gM46V999.gpGraLog(1, "Erro atualizar dados de contato - " & lMensagem)
        Unload Me
    End If

    lSql = "UPDATE      TTI" & vbCrLf
    lSql = lSql & "SET         PESS_CONT = " & gM46V999.gfFormatarCamposSQL(gM46V999.gfRetirarCaracterEspecial(Trim(txtContato(1).Text), False, True, True)) & "," & vbCrLf
    lSql = lSql & "            NUM_TEL   = " & gM46V999.gfFormatarCamposSQL(lTelefone) & "" & vbCrLf
    lSql = lSql & "FROM        RamosDiversos.dbo.Tab_Trans_Insp  TTI" & vbCrLf
    lSql = lSql & "INNER JOIN  RamosDiversos.dbo.Tab_Ped_Loc     TPL with (NOLOCK)" & vbCrLf
    lSql = lSql & "        ON  TPL.NOSSO_NUMERO                = TTI.NOSSO_NUMERO" & vbCrLf
    lSql = lSql & "       AND  TPL.NUM_ITEM                    = TTI.NUM_ITEM" & vbCrLf
    lSql = lSql & "WHERE       TPL.NOSSO_NUMERO                IN (" & TodosNossoNumero & ")" & vbCrLf
    
    If gM46V999.gfExeSQL(bdBasDados, lSql, lMensagem) <> 0 Then
        Call gM46V999.gpGraLog(1, "Erro atualizar dados de contato - " & lMensagem)
        Unload Me
    End If
    
    
    Dim lreg As Long
    
    '' 2016-04-12: Juliano, simplificando o processo que remove dados de contato (similar form T46v101A)
    lSql = "DELETE FROM  RamosDiversos.dbo.Tab_Ped_Insp_Compl " & vbCrLf
    lSql = lSql & "WHERE NOSSO_NUMERO IN (" & TodosNossoNumero & ")" & vbCrLf
    
    If gM46V999.gfExeSQL(bdBasDados, lSql, lMensagem, lreg) <> 0 Then
        Call gM46V999.gpGraLog(1, "Erro atualizar dados de contato - " & lMensagem)
        Unload Me
    End If
  
    '' 2016-04-12: Juliano, e os replica
    lSql = "INSERT INTO RamosDiversos.dbo.Tab_Ped_Insp_Compl " & vbCrLf
    lSql = lSql & "SELECT       TPL.Cod_Ramo, " & vbCrLf
    lSql = lSql & "             TPL.Nosso_Numero, " & vbCrLf
    lSql = lSql & "             TPL.NUM_ITEM, " & vbCrLf
    lSql = lSql & "             TPL.NUM_PI, " & vbCrLf
    lSql = lSql & "             " & gM46V999.gfFormatarCamposSQL(gM46V999.gfRetirarCaracterEspecial(Trim(txtContato(1).Text), False, True, True)) & ", " & vbCrLf
    lSql = lSql & "             " & gM46V999.gfFormatarCamposSQL(gM46V999.gfRetirarCaracterEspecial(Trim(txtContato(2).Text), False, True, True)) & ", " & vbCrLf
    lSql = lSql & "             " & gM46V999.gfFormatarCamposSQL(Trim(txtTelefone(1).Text)) & ", " & vbCrLf
    lSql = lSql & "             " & gM46V999.gfFormatarCamposSQL(Trim(txtTelefone(2).Text)) & ", " & vbCrLf
    lSql = lSql & "             " & gM46V999.gfFormatarCamposSQL(txtEmail.Text) & ", " & vbCrLf
    lSql = lSql & "             " & gM46V999.gfFormatarCamposSQL(txtPontoReferencia.Text) & vbCrLf
    lSql = lSql & "FROM        RamosDiversos.dbo.Tab_Ped_Loc TPL" & vbCrLf
    lSql = lSql & "WHERE       TPL.NOSSO_NUMERO IN (" & TodosNossoNumero & ")" & vbCrLf
                
    If gM46V999.gfExeSQL(bdBasDados, lSql, lMensagem, lreg) <> 0 Then
        Call gM46V999.gpGraLog(1, "Erro atualizar dados de contato - " & lMensagem)
        Unload Me
    End If
    
    Screen.MousePointer = vbDefault
        
    MsgBox "Operação concluida com sucesso!" & vbCrLf & _
           "Replicados " & lreg / qtdNN(TodosNossoNumero) & _
           " registros", vbInformation, Me.Caption

    TodosItensCorretos = True
    Unload Me
End Sub

Private Function qtdNN(strNN As String) As Integer
    If (strNN <> "" And InStr(1, strNN, ",") > 0) Then
        Dim arrNN As Variant
        arrNN = Split(strNN, ",")
        qtdNN = UBound(arrNN) + 1
    Else
        qtdNN = 1
    End If
End Function

Private Sub cmdSalvar_Click()
    Dim Num_Item As Long
    Dim lNome As String
    Dim lTelefone As String
    Dim lSql As String
    Dim lrst As clsYasRecordSet
    Dim bdBasDados As ADODB.Connection    'Base de dados.
    Dim lMensagem As String
    
    Screen.MousePointer = vbHourglass
    
    'Telefone 1
    lTelefone = Trim(txtTelefone(0).Text)
    If gM46V999.gfPreenchido(lTelefone) = False Then
        Call gM46V999.gpGraLog(1, "FGE0105 - Telefone 1 de contato para inspeção inválido (não preenchido).")
        Screen.MousePointer = vbDefault
        Exit Sub
        
    Else
        '1000477 - Telefone 1 de Contato Formatado
        '1 - valida se esta totalmente preenchido
        If Len(lTelefone) < 13 Then
            Call gM46V999.gpGraLog(1, "FGE0105 - Telefone 1 de contato para inspeção inválido (Favor informar o DDD e o número do Telefone 1 com 8 dígitos.)" & vbCrLf & "Formato: (99)9999-9999.")
            Screen.MousePointer = vbDefault
            Exit Sub
            
        Else
            '2 - valida formado (99)9999-9999
            If Mid(lTelefone, 1, 1) <> "(" Or Mid(lTelefone, 4, 1) <> ")" Or (Mid(lTelefone, 9, 1) <> "-" And Mid(lTelefone, 10, 1) <> "-") Then    '1200178 - Telefone 1 9 Digitos
                Call gM46V999.gpGraLog(1, "FGE0105 - Telefone 1 de contato para inspeção inválido (Favor informar o DDD e o número do Telefone 1 com 8 dígitos.)" & vbCrLf & "Formato: (99)9999-9999.")
                Screen.MousePointer = vbDefault
                Exit Sub
            Else
                '3 - não permite todos os numeros do Telefone 1 iguais
                If (InStr(1, "0000-0000 1111-1111 2222-2222 3333-3333 4444-4444 5555-5555 6666-6666 7777-7777 8888-8888 9999-9999", Mid(lTelefone, 5, Len(lTelefone)), vbTextCompare) > 0) Or _
                   (InStr(1, "00000-0000 11111-1111 22222-2222 33333-3333 44444-4444 55555-5555 66666-6666 77777-7777 88888-8888 99999-9999", Mid(lTelefone, 5, Len(lTelefone)), vbTextCompare) > 0) Then   ''1200178 - Telefone 1 9 Digitos
                    Call gM46V999.gpGraLog(1, "FGE0105 - Telefone 1 de contato para inspeção inválido (Favor informar o DDD e o número do Telefone 1 com 8 dígitos.)" & vbCrLf & "Formato: (99)99999-9999.")
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            End If
        End If
    End If
    
    'Telefone2
    lTelefone = Trim(txtTelefone(1).Text)
    If gM46V999.gfPreenchido(lTelefone) = False Then
        Call gM46V999.gpGraLog(1, "FGE0105 - Telefone 2 de contato para inspeção inválido (não preenchido).")
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        '1000477 - Telefone 2 de Contato Formatado
        '1 - valida se esta totalmente preenchido
        If Len(lTelefone) < 13 Then
            Call gM46V999.gpGraLog(1, "FGE0105 - Telefone 2 de contato para inspeção inválido (Favor informar o DDD e o número do Telefone 2 com 8 dígitos.)" & vbCrLf & "Formato: (99)9999-9999.")
            Screen.MousePointer = vbDefault
            Exit Sub
        Else
            '2 - valida formado (99)9999-9999
            If Mid(lTelefone, 1, 1) <> "(" Or Mid(lTelefone, 4, 1) <> ")" Or (Mid(lTelefone, 9, 1) <> "-" And Mid(lTelefone, 10, 1) <> "-") Then    '1200178 - Telefone 2 9 Digitos
                Call gM46V999.gpGraLog(1, "FGE0105 - Telefone 2 de contato para inspeção inválido (Favor informar o DDD e o número do Telefone 2 com 8 dígitos.)" & vbCrLf & "Formato: (99)9999-9999.")
                Screen.MousePointer = vbDefault
                Exit Sub
            Else
                '3 - não permite todos os numeros do Telefone 2 iguais
                If (InStr(1, "0000-0000 1111-1111 2222-2222 3333-3333 4444-4444 5555-5555 6666-6666 7777-7777 8888-8888 9999-9999", Mid(lTelefone, 5, Len(lTelefone)), vbTextCompare) > 0) Or _
                   (InStr(1, "00000-0000 11111-1111 22222-2222 33333-3333 44444-4444 55555-5555 66666-6666 77777-7777 88888-8888 99999-9999", Mid(lTelefone, 5, Len(lTelefone)), vbTextCompare) > 0) Then   ''1200178 - Telefone 2 9 Digitos
                    Call gM46V999.gpGraLog(1, "FGE0105 - Telefone 2 de contato para inspeção inválido (Favor informar o DDD e o número do Telefone 2 com 8 dígitos.)" & vbCrLf & "Formato: (99)99999-9999.")
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            End If
        End If
    End If
    
    'Telefone3 não é obrigatorio
    lTelefone = Trim(txtTelefone(2).Text)
    If gM46V999.gfPreenchido(lTelefone) Then
        '1000477 - Telefone 3 de Contato Formatado
        '1 - valida se esta totalmente preenchido
        If Len(lTelefone) < 13 Then
            Call gM46V999.gpGraLog(1, "FGE0105 - Telefone 3 de contato para inspeção inválido (Favor informar o DDD e o número do Telefone 3 com 8 dígitos.)" & vbCrLf & "Formato: (99)9999-9999.")
            Screen.MousePointer = vbDefault
            Exit Sub
        Else
            '2 - valida formado (99)9999-9999
            If Mid(lTelefone, 1, 1) <> "(" Or Mid(lTelefone, 4, 1) <> ")" Or (Mid(lTelefone, 9, 1) <> "-" And Mid(lTelefone, 10, 1) <> "-") Then    '1200178 - Telefone 3 9 Digitos
                Call gM46V999.gpGraLog(1, "FGE0105 - Telefone 3 de contato para inspeção inválido (Favor informar o DDD e o número do Telefone 3 com 8 dígitos.)" & vbCrLf & "Formato: (99)9999-9999.")
                Screen.MousePointer = vbDefault
                Exit Sub
            Else
                '3 - não permite todos os numeros do Telefone 3 iguais
                If (InStr(1, "0000-0000 1111-1111 2222-2222 3333-3333 4444-4444 5555-5555 6666-6666 7777-7777 8888-8888 9999-9999", Mid(lTelefone, 5, Len(lTelefone)), vbTextCompare) > 0) Or _
                   (InStr(1, "00000-0000 11111-1111 22222-2222 33333-3333 44444-4444 55555-5555 66666-6666 77777-7777 88888-8888 99999-9999", Mid(lTelefone, 5, Len(lTelefone)), vbTextCompare) > 0) Then   ''1200178 - Telefone 3 9 Digitos
                    Call gM46V999.gpGraLog(1, "FGE0105 - Telefone 3 de contato para inspeção inválido (Favor informar o DDD e o número do Telefone 3 com 8 dígitos.)" & vbCrLf & "Formato: (99)99999-9999.")
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            End If
        End If
    End If

    If (gM46V999.gfPreenchido(txtContato(0).Text) = False Or Len(txtContato(0).Text) <= 3) Or _
       (gM46V999.gfPreenchido(txtContato(1).Text) = False Or Len(txtContato(1).Text) <= 3) Then
        Call gM46V999.gpGraLog(1, "FGE0106 - Necessario informar ao menos dois contatos para inspeção válidos (informar no mínimo 3 caracteres).")
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    ' Salvar os dados
    If gM46V999.gfAbrBasDados("P00MPED", gM46V999.gTipP00MPED, gM46V999.gEndP00MPED, bdBasDados, lMensagem) = False Then
        Call gM46V999.gpGraLog(1, "Erro ao abrir base de dados - " & lMensagem)
        Unload Me
    End If
    
    ' Obter os dados digitados
    Num_Item = CLng(lblNumItem.Caption)
    lNome = gM46V999.gfRetirarCaracterEspecial(Trim(txtContato(0).Text), False, True, True)
    lTelefone = Trim(txtTelefone(0).Text)
    
    lSql = "UPDATE      TPI" & vbCrLf
    lSql = lSql & "SET         PESS_CONT  = " & gM46V999.gfFormatarCamposSQL(lNome) & "," & vbCrLf
    lSql = lSql & "            NUM_TEL = " & gM46V999.gfFormatarCamposSQL(lTelefone) & "" & vbCrLf
    lSql = lSql & "FROM        RamosDiversos.dbo.Tab_Ped_Insp  TPI" & vbCrLf
    lSql = lSql & "INNER JOIN  RamosDiversos.dbo.Tab_Ped_Loc   TPL with (NOLOCK)" & vbCrLf
    lSql = lSql & "    ON      TPL.NOSSO_NUMERO                = TPI.NOSSO_NUMERO" & vbCrLf
    lSql = lSql & "        AND TPL.NUM_ITEM                    = TPI.NUM_ITEM" & vbCrLf
    lSql = lSql & "WHERE       TPL.NOSSO_NUMERO                IN (" & TodosNossoNumero & ")" & vbCrLf
    lSql = lSql & "AND         TPL.NUM_ITEM                    = " & Num_Item & vbCrLf
    
    If gM46V999.gfExeSQL(bdBasDados, lSql, lMensagem) <> 0 Then
        Call gM46V999.gpGraLog(1, "Erro atualizar dados de contato - " & lMensagem)
        Unload Me
    End If

    lSql = "UPDATE      TPI" & vbCrLf
    lSql = lSql & "SET         PESS_CONT  = " & gM46V999.gfFormatarCamposSQL(gM46V999.gfRetirarCaracterEspecial(Trim(txtContato(1).Text), False, True, True)) & "," & vbCrLf
    lSql = lSql & "            NUM_TEL = " & gM46V999.gfFormatarCamposSQL(lTelefone) & "" & vbCrLf
    lSql = lSql & "FROM        RamosDiversos.dbo.Tab_Trans_Insp  TPI" & vbCrLf
    lSql = lSql & "INNER JOIN  RamosDiversos.dbo.Tab_Ped_Loc   TPL with (NOLOCK)" & vbCrLf
    lSql = lSql & "    ON      TPL.NOSSO_NUMERO                = TPI.NOSSO_NUMERO" & vbCrLf
    lSql = lSql & "        AND TPL.NUM_ITEM                    = TPI.NUM_ITEM" & vbCrLf
    lSql = lSql & "WHERE       TPL.NOSSO_NUMERO                IN (" & TodosNossoNumero & ")" & vbCrLf
    lSql = lSql & "AND         TPL.NUM_ITEM                    = " & Num_Item & vbCrLf
    
    If gM46V999.gfExeSQL(bdBasDados, lSql, lMensagem) <> 0 Then
        Call gM46V999.gpGraLog(1, "Erro atualizar dados de contato - " & lMensagem)
        Unload Me
    End If
    
    Dim lreg As Long
    
    lSql = "UPDATE      TPI" & vbCrLf
    lSql = lSql & "SET         PESS_CONTATO_2  = " & gM46V999.gfFormatarCamposSQL(gM46V999.gfRetirarCaracterEspecial(Trim(txtContato(1).Text), False, True, True)) & "," & vbCrLf
    lSql = lSql & "            PESS_CONTATO_3  = " & gM46V999.gfFormatarCamposSQL(gM46V999.gfRetirarCaracterEspecial(Trim(txtContato(2).Text), False, True, True)) & "," & vbCrLf
    lSql = lSql & "            NUM_TEL_2 = " & gM46V999.gfFormatarCamposSQL(Trim(txtTelefone(1).Text)) & ", " & vbCrLf
    lSql = lSql & "            NUM_TEL_3 = " & gM46V999.gfFormatarCamposSQL(Trim(txtTelefone(2).Text)) & ", " & vbCrLf
    lSql = lSql & "            DSC_EMAIL = " & gM46V999.gfFormatarCamposSQL(txtEmail.Text) & ", " & vbCrLf
    lSql = lSql & "            DSC_PONTO_REF = " & gM46V999.gfFormatarCamposSQL(txtPontoReferencia.Text) & " " & vbCrLf
    lSql = lSql & "FROM        RamosDiversos.dbo.Tab_Ped_Insp_Compl  TPI " & vbCrLf
    lSql = lSql & "INNER JOIN  RamosDiversos.dbo.Tab_Ped_Loc   TPL with (NOLOCK)" & vbCrLf
    lSql = lSql & "    ON      TPL.NOSSO_NUMERO                = TPI.NOSSO_NUMERO" & vbCrLf
    lSql = lSql & "        AND TPL.NUM_ITEM                    = TPI.NUM_ITEM" & vbCrLf
    lSql = lSql & "WHERE       TPL.NOSSO_NUMERO                IN (" & TodosNossoNumero & ")" & vbCrLf
    lSql = lSql & "AND         TPL.NUM_ITEM                    = " & Num_Item & vbCrLf
    
    If gM46V999.gfExeSQL(bdBasDados, lSql, lMensagem, lreg) <> 0 Then
        Call gM46V999.gpGraLog(1, "Erro atualizar dados de contato - " & lMensagem)
        Unload Me
    End If
    
    If lreg = 0 Then
        lSql = "INSERT INTO RamosDiversos.dbo.Tab_Ped_Insp_Compl " & vbCrLf
        lSql = lSql & "SELECT       TPL.Cod_Ramo, " & vbCrLf
        lSql = lSql & "             TPL.Nosso_Numero, " & vbCrLf
        lSql = lSql & "             TPL.NUM_ITEM, " & vbCrLf
        lSql = lSql & "             TPL.NUM_PI, " & vbCrLf
        lSql = lSql & "             " & gM46V999.gfFormatarCamposSQL(gM46V999.gfRetirarCaracterEspecial(Trim(txtContato(1).Text), False, True, True)) & ", " & vbCrLf
        lSql = lSql & "             " & gM46V999.gfFormatarCamposSQL(gM46V999.gfRetirarCaracterEspecial(Trim(txtContato(2).Text), False, True, True)) & ", " & vbCrLf
        lSql = lSql & "             " & gM46V999.gfFormatarCamposSQL(Trim(txtTelefone(1).Text)) & ", " & vbCrLf
        lSql = lSql & "             " & gM46V999.gfFormatarCamposSQL(Trim(txtTelefone(2).Text)) & ", " & vbCrLf
        lSql = lSql & "             " & gM46V999.gfFormatarCamposSQL(txtEmail.Text) & ", " & vbCrLf
        lSql = lSql & "             " & gM46V999.gfFormatarCamposSQL(txtPontoReferencia.Text) & vbCrLf
        lSql = lSql & "FROM        RamosDiversos.dbo.Tab_Ped_Loc   TPL" & vbCrLf
        lSql = lSql & " " & vbCrLf
        lSql = lSql & "    " & vbCrLf
        lSql = lSql & "        " & vbCrLf
        lSql = lSql & "WHERE       TPL.NOSSO_NUMERO                IN (" & TodosNossoNumero & ")" & vbCrLf
        lSql = lSql & "AND         TPL.NUM_ITEM                    = " & Num_Item & vbCrLf
                    
        If gM46V999.gfExeSQL(bdBasDados, lSql, lMensagem, lreg) <> 0 Then
            Call gM46V999.gpGraLog(1, "Erro atualizar dados de contato - " & lMensagem)
            Unload Me
        End If
    End If
    
    Screen.MousePointer = vbDefault
    
    ' Chamar próximo item
    VerificarItens Num_Item
End Sub

Private Sub Form_Activate()
    Me.AutoRedraw = False
    TodosItensCorretos = False
    PonteiroMouse = Screen.MousePointer
    ConstroiStringNossoNumero
    VerificarItens 0
    
    Screen.MousePointer = vbDefault
    
    ' Se chegou aqui com todos os itens corretos, não precisa mostrar o form
    If TodosItensCorretos Then
        Unload Me
    
    Else
        Me.AutoRedraw = True
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = PonteiroMouse
End Sub


Private Sub txtTelefone_LostFocus(Index As Integer)
    With txtTelefone(Index)
        If InStr(1, .Text, "-", vbTextCompare) > 0 Then Exit Sub
        Select Case Len(.Text)
        Case 12    '(99)9999-9999
            .Text = Left(.Text, 8) & "-" & Right(.Text, 4)
        Case 13    '(99)99999-9999
            .Text = Left(.Text, 9) & "-" & Right(.Text, 4)
        End Select
    End With
End Sub

Private Sub txtTelefone_Change(Index As Integer)
    Call gpLimPremios
    '1000477 - Telefone de Contato Formatado
    If Tecla = vbKeyBack Then Exit Sub
    
    With txtTelefone(Index)
        Select Case Len(.Text)
        Case 1
            If InStr(1, .Text, "(", vbTextCompare) = 0 Then
                .Text = "(" & .Text
                .SelStart = Len(.Text)
            End If
        Case 3
            If InStr(1, .Text, ")", vbTextCompare) = 0 Then
                .Text = .Text & ")"
                .SelStart = Len(.Text)
            End If
        End Select
    End With
End Sub

Private Sub txtTelefone_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    '1000477 - Telefone de Contato Formatado
    With txtTelefone(Index)
        Select Case KeyCode
        Case vbKeyDelete
            If .SelLength <> Len(.Text) Then
                Select Case Mid(.Text, .SelStart + 1, 1)
                Case "(", ")", "-"
                    KeyCode = 0
                End Select
            End If
        Case vbKeyBack
            If .SelStart > 0 And .SelStart <> Len(.Text) Then
                Select Case Mid(.Text, .SelStart, 1)
                Case "(", ")", "-"
                    KeyCode = 0
                    .SelStart = .SelStart - 1
                End Select
            End If
        End Select
    End With
End Sub

Private Sub txtTelefone_GotFocus(Index As Integer)
    With txtTelefone(Index)
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtTelefone_KeyPress(Index As Integer, KeyAscii As Integer)
'1000477 - Telefone de Contato Formatado
    Tecla = KeyAscii
    
    With txtTelefone(Index)
        Select Case KeyAscii
        Case 48 To 57   'permite de 0 a 9
            'limita tamanho do campo
            If Len(.Text) >= 14 And .SelLength <> Len(.Text) Then KeyAscii = 0: Exit Sub   '1200178 - Telefone 9 Digitos
            'se foco estiver no inicio e o campo ainda nao foi todo preenchido
            If .SelStart <= 3 And Len(.Text) > 0 And .SelLength <> Len(.Text) Then
                If .SelStart <= 3 And Mid(.Text, 4, 1) = ")" Then
                    KeyAscii = 0: Exit Sub
                End If
                .Text = Mid(.Text, 1, IIf(.SelStart = 0, 1, .SelStart)) & Chr(KeyAscii) & Mid(.Text, IIf(.SelStart = 0, 1, .SelStart) + 1, Len(.Text))
                .SelStart = IIf(.SelStart <= 1, .SelStart + 2, Len(.Text))
                KeyAscii = 0
            End If
            
        Case vbKeyReturn
            KeyAscii = 0
            SendKeys "{TAB}"
        Case vbKeyBack  'permite back
        Case Else   'nao permite demais teclas
            KeyAscii = 0
        End Select
    End With
End Sub

Private Sub txtContato_Change(Index As Integer)
    Call gpLimPremios
End Sub

Private Sub txtContato_GotFocus(Index As Integer)
    With txtContato(Index)
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtContato_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtContato_LostFocus(Index As Integer)
    txtContato(Index).Text = gM46V999.gfRetirarCaracterEspecial(txtContato(Index).Text, True, False, True)
End Sub

Private Sub VerificarItens(Num_Item_Anterior As Long)
    Dim lSql As String
    Dim lrst As clsYasRecordSet
    Dim bdBasDados As ADODB.Connection    'Base de dados.
    Dim lMensagem As String
    
    Screen.MousePointer = vbHourglass
    
    If gM46V999.gfAbrBasDados("P00MPED", gM46V999.gTipP00MPED, gM46V999.gEndP00MPED, bdBasDados, _
                              lMensagem) = False Then
        Call gM46V999.gpGraLog(1, "Erro ao abrir base de dados - " & lMensagem)
        Unload Me
    End If
    
    ' Consultar base de dados usando o NossoNumeroOld
    lSql = "SELECT      TPL.Num_Item," & vbCrLf
    lSql = lSql & "            TP.NOM_SEGURADO," & vbCrLf
    lSql = lSql & "            TPL.INSP_OBR," & vbCrLf
    lSql = lSql & "            TPL.NOM_LOGR_RISCO," & vbCrLf
    lSql = lSql & "            TPI.NUM_LOGR_RISCO," & vbCrLf
    lSql = lSql & "            TPI.PESS_CONT," & vbCrLf
    lSql = lSql & "            TPI.Num_Tel," & vbCrLf
    lSql = lSql & "            TPI2.PESS_CONTATO_2," & vbCrLf
    lSql = lSql & "            TPL.DSC_COMPL_RISCO," & vbCrLf
    lSql = lSql & "            TPL.NOM_CID_RISCO," & vbCrLf
    lSql = lSql & "            TPL.SIG_UF_RISCO," & vbCrLf
    lSql = lSql & "            TPL.Num_Cep_Risco " & vbCrLf
    lSql = lSql & "FROM        RamosDiversos.dbo.Tab_Ped_Loc   TPL with (NOLOCK)" & vbCrLf
    lSql = lSql & "INNER JOIN  RamosDiversos.dbo.Tab_Ped       TP  with (NOLOCK)" & vbCrLf
    lSql = lSql & "    ON      TP.NOSSO_NUMERO                 = TPL.NOSSO_NUMERO" & vbCrLf
    lSql = lSql & "LEFT JOIN   RamosDiversos.dbo.Tab_Ped_Loc_Compl   TPLC with (NOLOCK)" & vbCrLf
    lSql = lSql & "    ON      TPLC.NOSSO_NUMERO = TPL.NOSSO_NUMERO " & vbCrLf
    lSql = lSql & "    AND      TPLC.Num_Item = TPL.Num_Item " & vbCrLf
    lSql = lSql & "INNER JOIN  RamosDiversos.dbo.Tab_Ped_Insp  TPI with (NOLOCK)" & vbCrLf
    lSql = lSql & "    ON      TPI.NOSSO_NUMERO                = TPL.NOSSO_NUMERO" & vbCrLf
    lSql = lSql & "        AND TPI.NUM_ITEM                    = TPL.NUM_ITEM" & vbCrLf
    lSql = lSql & "LEFT JOIN  RamosDiversos.dbo.Tab_Ped_Insp_Compl  TPI2 with (NOLOCK)" & vbCrLf
    lSql = lSql & "    ON      TPI2.NOSSO_NUMERO                = TPL.NOSSO_NUMERO" & vbCrLf
    lSql = lSql & "        AND TPI2.NUM_ITEM                    = TPL.NUM_ITEM" & vbCrLf
    lSql = lSql & "WHERE       TPL.NOSSO_NUMERO                = '" & NossoNumeroOld & "'" & vbCrLf
    lSql = lSql & "        AND TPL.NUM_ITEM                    > " & Num_Item_Anterior & vbCrLf
    lSql = lSql & "        AND NOT TPLC.Ind_StatusInspecao IN (1,3) " & vbCrLf
    lSql = lSql & "ORDER BY    TPL.NUM_ITEM" & vbCrLf
    
    If Not gM46V999.gfObtRegistro(bdBasDados, lSql, lrst, lMensagem) Then
        Call gM46V999.gpGraLog(1, "Erro ao obter itens - " & lMensagem)
        Unload Me ' Fechar a tela... não faz sentido manter aberta.
    End If
    
    If lrst.EOF Then
        TodosItensCorretos = True
        Unload Me ' Fechar a tela... não faz sentido manter aberta.
    End If
    
    Do Until lrst.EOF
        If lrst("Num_Item") > Num_Item_Anterior Then
            If lrst("INSP_OBR") = 1 Then
                If Not (gM46V999.gfPreenchido(lrst("PESS_CONT")) And gM46V999.gfPreenchido(lrst("PESS_CONTATO_2"))) Then
                    ' Popular a tela
                    lblNumItem.Caption = lrst("Num_Item")
                    lblNomeSegurado.Caption = Trim(lrst("NOM_SEGURADO"))
                    lblEndereco.Caption = Trim(lrst("NOM_LOGR_RISCO"))
                    lblNumEndereco.Caption = lrst("NUM_LOGR_RISCO")
                    lblComplemento.Caption = Trim(lrst("DSC_COMPL_RISCO"))
                    lblCidade.Caption = Trim(lrst("NOM_CID_RISCO"))
                    lblUF.Caption = Trim(lrst("SIG_UF_RISCO"))
                    lblCEP.Caption = Mid$(Format(lrst("Num_Cep_Risco"), "00000000"), 1, 5) & "-" & Mid$(Format(lrst("Num_Cep_Risco"), "00000000"), 6, 3)
                    ' Sair da Sub sem marcar o final
                    
                    If gM46V999.gfPreenchido(lrst("PESS_CONT")) Then
                        txtContato(0).Text = gM46V999.gfTratarNulo(lrst("PESS_CONT"), tpoString)
                        txtTelefone(0).Text = gM46V999.gfTratarNulo(lrst("Num_Tel"), tpoString)
                        txtContato(1).SetFocus
                    Else
                        txtContato(0).SetFocus
                    End If
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            End If
        End If
        lrst.MoveNext
    Loop
    
    ' Chegou aqui é porque todos os itens estão devidamente preenchidos.
    TodosItensCorretos = True
    Unload Me
    
End Sub

Private Sub ConstroiStringNossoNumero()
    Dim lSql As String
    Dim lrst As clsYasRecordSet
    Dim bdBasDados As ADODB.Connection    'Base de dados.
    Dim lMensagem As String
    
    TodosNossoNumero = "'" & NossoNumeroOld & "'"
    
    If gM46V999.gfAbrBasDados("P00MPED", gM46V999.gTipP00MPED, gM46V999.gEndP00MPED, bdBasDados, _
                              lMensagem) = False Then
        Call gM46V999.gpGraLog(1, "Erro ao abrir base de dados - " & lMensagem)
        Unload Me
    End If
    
    lSql = "SELECT TPNC.NOSSO_NUMERO_NEW as NOSSO_NUMERO FROM RamosDiversos..Tab_Ped_Num_Copia TPNC WHERE NOSSO_NUMERO_OLD = '" & NossoNumeroOld & "'" & vbCrLf
    If Not gM46V999.gfObtRegistro(bdBasDados, lSql, lrst, lMensagem) Then
        Call gM46V999.gpGraLog(1, "Erro ao obter itens - " & lMensagem)
        Unload Me ' Fechar a tela... não faz sentido manter aberta.
    End If
    
    Do Until lrst.EOF
        TodosNossoNumero = TodosNossoNumero & ", '" & Trim(lrst("NOSSO_NUMERO")) & "'"
        lrst.MoveNext
    Loop
    
    lSql = "SELECT TPNC.NOSSO_NUMERO_OLD as NOSSO_NUMERO FROM RamosDiversos..Tab_Ped_Num_Copia TPNC WHERE NOSSO_NUMERO_NEW = '" & NossoNumeroOld & "'" & vbCrLf
    If Not gM46V999.gfObtRegistro(bdBasDados, lSql, lrst, lMensagem) Then
        Call gM46V999.gpGraLog(1, "Erro ao obter itens - " & lMensagem)
        Unload Me ' Fechar a tela... não faz sentido manter aberta.
    End If
    
    Do Until lrst.EOF
        TodosNossoNumero = TodosNossoNumero & ", '" & Trim(lrst("NOSSO_NUMERO")) & "'"
        lrst.MoveNext
    Loop
    
End Sub
