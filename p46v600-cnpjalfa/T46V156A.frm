VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmT46V156A 
   Caption         =   "Verificação de Resultados"
   ClientHeight    =   3750
   ClientLeft      =   315
   ClientTop       =   1620
   ClientWidth     =   12735
   Icon            =   "T46V156A.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3750
   ScaleWidth      =   12735
   Begin Threed.SSCommand cmdOk 
      Height          =   300
      Left            =   5205
      TabIndex        =   1
      ToolTipText     =   "Volta à proposta."
      Top             =   3375
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Ok"
      ForeColor       =   8388608
   End
   Begin Threed.SSCommand cmdSalvar 
      Height          =   300
      Left            =   6300
      TabIndex        =   2
      ToolTipText     =   "Volta à proposta."
      Top             =   3375
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   529
      _StockProps     =   78
      Caption         =   "&Salvar"
      ForeColor       =   8388608
   End
   Begin MSFlexGridLib.MSFlexGrid flxResult 
      Height          =   3375
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   12750
      _ExtentX        =   22490
      _ExtentY        =   5953
      _Version        =   393216
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
End
Attribute VB_Name = "frmT46V156A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const ColNovoPremio As Integer = 8
Const ColPremioCalculado As Integer = 3

Private Sub cmdOk_Click()
    Unload Me
End Sub


Private Sub cmdSalvar_Click()
    Dim lobjstcPedCobert As stcA46V704B
    Set lobjstcPedCobert = New stcA46V704B
    Set lobjstcPedCobert.M46V999 = gM46V999
    Dim ifx As Integer

    For Each lobjstcPedCobert In frmT46V101A.objstcPedLoc.ColecaoPedCobert
        For ifx = 1 To flxResult.Rows - 1
            If flxResult.TextMatrix(ifx, 0) = lobjstcPedCobert.Cod_Cobert Then
                If flxResult.TextMatrix(ifx, 8) = "" Then
                    ' Não quer prêmio informado - Manter o original
                    If lobjstcPedCobert.PremioInformado = True Then
                        ' Tinha Premio informado e agora precisa voltar o original - se não entrar nesse if, não precisa fazer nada.
                        lobjstcPedCobert.VAL_PR = lobjstcPedCobert.Val_Pr_Orig
                        lobjstcPedCobert.PremioInformado = False
                    End If
                Else
                    ' Alterou o prêmio
                    lobjstcPedCobert.Val_Pr_Orig = lobjstcPedCobert.VAL_PR
                    lobjstcPedCobert.PremioInformado = True
                    lobjstcPedCobert.VAL_PR = flxResult.TextMatrix(ifx, 8)
                End If

            End If
        Next
    Next

    ' 2017-12-21: Obrigar calcular antes de Reanálise
    frmT46V101A.flagCalcular = True
    
    ' Invocar um novo cálculo

    ' Fechar a janela
    Call cmdOk_Click

End Sub

Private Sub flxResult_KeyDown(KeyCode As Integer, Shift As Integer)
    With flxResult
        .Row = .RowSel
        If .Row > 0 Then
            If .ColSel = ColNovoPremio Then
                Select Case KeyCode
                Case 48 To 57, 97, 101, 109
                    .TextMatrix(.RowSel, .ColSel) = .TextMatrix(.RowSel, .ColSel)    ' & Chr(KeyCode)
                Case 46
                    .TextMatrix(.RowSel, .ColSel) = ""
                End Select
            End If
        End If
    End With

End Sub


Private Sub flxResult_KeyPress(KeyAscii As Integer)

    With flxResult
        If .RowSel = 0 Then Exit Sub
        Select Case KeyAscii
        Case 8
            If Len(.TextMatrix(.RowSel, .ColSel)) > 0 Then
                .TextMatrix(.RowSel, .ColSel) = Left(.TextMatrix(.RowSel, .ColSel), Len(.TextMatrix(.RowSel, .ColSel)) - 1)
            End If
        Case 13
            If IsNumeric(.TextMatrix(.RowSel, .ColSel)) Then
                flxResult.TextMatrix(.RowSel, .ColSel) = Format(.TextMatrix(.RowSel, .ColSel), "standard")
            End If
        Case 44
            If InStr(1, .TextMatrix(.RowSel, .ColSel), Chr(44)) > 0 Then Exit Sub
            .TextMatrix(.RowSel, .ColSel) = .TextMatrix(.RowSel, .ColSel) & Chr(KeyAscii)
        Case 48 To 57, 45
            .TextMatrix(.RowSel, .ColSel) = .TextMatrix(.RowSel, .ColSel) & Chr(KeyAscii)
        End Select
    End With

End Sub


Private Sub Form_Load()
    Dim lobjstcPedCobert As stcA46V704B
    Dim lI As Integer
    Dim Senha As String            'Mesma do banco


    If App.EXEName <> "P46V101" Then
        If frmT46V101A.barFerramentas.Buttons(24).Tag = "" Then
            Senha = InputBox("Entre com a Senha")
            If Senha <> "p00multi46" Then Exit Sub
        End If
        frmT46V101A.barFerramentas.Buttons(24).Tag = Senha
    End If

    '1. Posiciona form.
    Call gM46V999.gpCenForm(Me)

    flxResult.Cols = 9

    flxResult.Row = 0
    flxResult.FixedCols = 0

    flxResult.Col = 0
    flxResult.Text = "Cobert"
    flxResult.Col = 1
    flxResult.Text = "Vlr LMG"
    flxResult.Col = 2
    flxResult.Text = "Valores da Fórmula"
    flxResult.Col = 3
    flxResult.Text = "Pr Calculado"
    flxResult.Col = 4
    flxResult.Text = "Pr Anual"
    flxResult.Col = 5
    flxResult.Text = "Pr Resseguro"
    flxResult.Col = 6
    flxResult.Text = "Pr Tarifário"
    flxResult.Col = 7
    flxResult.Text = "Pr Original"
    flxResult.Col = 8
    flxResult.Text = "Novo Prêmio"

    flxResult.ColWidth(0) = 500
    flxResult.ColWidth(1) = 1200
    flxResult.ColWidth(2) = 5700
    flxResult.ColWidth(3) = 800
    flxResult.ColWidth(4) = 800
    flxResult.ColWidth(5) = 800
    flxResult.ColWidth(6) = 800
    flxResult.ColWidth(7) = 800
    flxResult.ColWidth(8) = 800

    Set lobjstcPedCobert = New stcA46V704B
    Set lobjstcPedCobert.M46V999 = gM46V999

    lI = 1
    For Each lobjstcPedCobert In frmT46V101A.objstcPedLoc.ColecaoPedCobert
        flxResult.TextMatrix(lI, 0) = Format$(lobjstcPedCobert.Cod_Cobert, "0000")
        flxResult.TextMatrix(lI, 1) = Format$(lobjstcPedCobert.Val_IS, "standard")
        flxResult.TextMatrix(lI, 2) = lobjstcPedCobert.Formula
        flxResult.TextMatrix(lI, 3) = Format$(lobjstcPedCobert.VAL_PR, "standard")
        flxResult.TextMatrix(lI, 4) = Format$(lobjstcPedCobert.Val_Pr_Anu, "standard")
        flxResult.TextMatrix(lI, 5) = Format$(lobjstcPedCobert.Val_Pr_Resg, "standard")
        flxResult.TextMatrix(lI, 6) = Format$(lobjstcPedCobert.Val_Pr_Tarif, "standard")
        flxResult.TextMatrix(lI, 7) = Format$(lobjstcPedCobert.Val_Pr_Orig, "standard")


        If lobjstcPedCobert.PremioInformado = True Then
            ' Se for premio informado, o valor do premio calculado está na propriedade Val_Pr_Orig
            ' Precisa colocar este valor no campo correto, assim como o valor do premio informado
            flxResult.TextMatrix(lI, 8) = Format$(lobjstcPedCobert.VAL_PR, "standard")
            flxResult.TextMatrix(lI, 3) = Format$(lobjstcPedCobert.Val_Pr_Orig, "standard")
        Else
            ' Não é Premio Informado, então a coluna de Valor de Premio Original deve ser igual a do valor do premio calculado
            flxResult.TextMatrix(lI, 7) = Format$(lobjstcPedCobert.VAL_PR, "standard")
        End If

        lI = lI + 1
        flxResult.Rows = lI + 1
    Next
    flxResult.Rows = lI
End Sub

Private Sub Form_Resize()

    On Error GoTo Erro
    
    If Me.Height > 1500 And Me.Width > 1500 Then
        With flxResult
            .Left = 30
            .Top = 30
            .Width = Me.Width - 270
            .Height = Me.Height - 1100
        End With
        
        cmdOk.Top = Me.Height - 1000
        cmdOk.Left = CInt((Me.Width - (cmdOk.Width + cmdSalvar.Width + 60)) / 2) - 270
        cmdSalvar.Top = cmdOk.Top
        cmdSalvar.Left = cmdOk.Left + cmdOk.Width + 60
    End If
    
    Exit Sub
Erro:
    Call gM46V999.gpGraLog(0, Me.Name + "_Form_Resize ERRO - " & Err & " " & Error)
End Sub
