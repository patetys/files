VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmT46V202A 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Prêmio Informado"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   9300
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDistribuir 
      Caption         =   "Distribuir prêmios"
      Height          =   375
      Left            =   2430
      TabIndex        =   8
      Top             =   4635
      Width           =   1590
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "Limpar"
      Height          =   375
      Left            =   45
      TabIndex        =   7
      Top             =   4635
      Width           =   1590
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "Salvar"
      Height          =   375
      Left            =   7515
      TabIndex        =   6
      Top             =   4635
      Width           =   1590
   End
   Begin MSFlexGridLib.MSFlexGrid flxCoberturas 
      Height          =   3390
      Left            =   90
      TabIndex        =   5
      Top             =   360
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   5980
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.TextBox txtValPremioInformado 
      Height          =   330
      Left            =   2340
      TabIndex        =   2
      Top             =   4230
      Width           =   1680
   End
   Begin VB.Label lblCobs 
      Caption         =   "Coberturas"
      Height          =   240
      Left            =   90
      TabIndex        =   4
      Top             =   90
      Width           =   1590
   End
   Begin VB.Label lblPremioLiquidoCalc 
      Caption         =   "R$ 1000,00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2025
      TabIndex        =   3
      Top             =   3915
      Width           =   1995
   End
   Begin VB.Label Label2 
      Caption         =   "Prêmio Líquido Informado: R$ "
      Height          =   285
      Left            =   90
      TabIndex        =   1
      Top             =   4275
      Width           =   2220
   End
   Begin VB.Label Label1 
      Caption         =   "Prêmio Líquido Calculado:"
      Height          =   285
      Left            =   90
      TabIndex        =   0
      Top             =   3915
      Width           =   1950
   End
End
Attribute VB_Name = "frmT46V202A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mobjstcPed As stcA46V702B
Public ConexaoBD As ADODB.Connection    ' Conexão com o Banco de Dados
Public UltimoErro As String           ' Descrição do último erro gerado
Public Val_Pr_Liq As Double
Public Val_Pr_Tot_Orig As Double
Public Flg_PremioLiquidoBateComSoma As Boolean
Public Flg_ClicouSalvar As Boolean

Public Flg_TemQueFecharForm As Boolean

Public Enum T46V202_Grid
    Item = 1
    Codigo = 2
    Descricao = 3
    Val_pr_original = 4
    VAL_PR_Informado = 5
End Enum

Public Sub CarregaForm()
    Flg_TemQueFecharForm = False
    ' Somar os premios das coberturas para obter o premio liquido do cálculo
    
    ' Carregar o FlxGrid com as coberturas
    ConfigGrid
    CarregaGrid
    
End Sub

Private Sub CarregaGrid()
    
    Dim pobjstcPedLoc As stcA46V708B
    Dim pobjstcPedCobert As stcA46V704B
    Dim l As Long
    Dim c As Long
    Dim Linha As Long
    
    Dim Val_Pr_Tot_Inf As Double
    Dim Flg_Tudo_Premio_Informado As Boolean
    
    ' Iniciando variaveis
    Val_Pr_Liq = 0
    Val_Pr_Tot_Orig = 0
    Val_Pr_Tot_Inf = 0
    Flg_Tudo_Premio_Informado = True
    
    With flxCoberturas
    .Rows = 1
    
    For l = 1 To mobjstcPed.ColecaoPedLoc.Count
        Set pobjstcPedLoc = mobjstcPed.ColecaoPedLoc(l)
        For c = 1 To pobjstcPedLoc.ColecaoPedCobert.Count
            Set pobjstcPedCobert = pobjstcPedLoc.ColecaoPedCobert(c)
            
            .Redraw = False
            .Rows = .Rows + 1
            
            .Row = .Rows - 1
            
            .TextMatrix(.Rows - 1, T46V202_Grid.Item) = pobjstcPedLoc.Num_Item
            .Col = T46V202_Grid.Item
            .CellAlignment = flexAlignRightCenter
            
            .TextMatrix(.Rows - 1, T46V202_Grid.Codigo) = pobjstcPedCobert.Cod_Cobert
            .Col = T46V202_Grid.Codigo
            .CellAlignment = flexAlignRightCenter
            
            .TextMatrix(.Rows - 1, T46V202_Grid.Descricao) = pobjstcPedCobert.gNomeCobertura
            .Col = T46V202_Grid.Descricao
            .CellAlignment = flexAlignLeftCenter
            
            If pobjstcPedCobert.PremioInformado = False Then
                .TextMatrix(.Rows - 1, T46V202_Grid.Val_pr_original) = Format(pobjstcPedCobert.VAL_PR, "###,###,##0.00")
                .Col = T46V202_Grid.Val_pr_original
                .CellAlignment = flexAlignRightCenter
                
                .TextMatrix(.Rows - 1, T46V202_Grid.VAL_PR_Informado) = ""
                .Col = T46V202_Grid.VAL_PR_Informado
                .CellAlignment = flexAlignRightCenter
                
                Flg_Tudo_Premio_Informado = False
                Val_Pr_Tot_Orig = Val_Pr_Tot_Orig + pobjstcPedCobert.VAL_PR
            Else
                .TextMatrix(.Rows - 1, T46V202_Grid.Val_pr_original) = Format(pobjstcPedCobert.Val_Pr_Orig, "###,###,##0.00")
                .Col = T46V202_Grid.Val_pr_original
                .CellAlignment = flexAlignRightCenter
                
                .TextMatrix(.Rows - 1, T46V202_Grid.VAL_PR_Informado) = Format(pobjstcPedCobert.VAL_PR, "###,###,##0.00")
                .Col = T46V202_Grid.VAL_PR_Informado
                .CellAlignment = flexAlignRightCenter
                Val_Pr_Tot_Inf = Val_Pr_Tot_Inf + pobjstcPedCobert.VAL_PR
                Val_Pr_Tot_Orig = Val_Pr_Tot_Orig + pobjstcPedCobert.Val_Pr_Orig
            End If
            
            Val_Pr_Liq = Val_Pr_Liq + pobjstcPedCobert.VAL_PR
            
            .Redraw = True
            
        Next c
    Next l
    End With
    
    If Val_Pr_Tot_Orig = 0 Then
        MsgBox "Não é possível usar o Premio informado quando o prêmio original está zerado." & vbCrLf & _
        "Favor calcular o prêmio antes de usar o Premio informado", vbExclamation, "Aviso"
        Flg_TemQueFecharForm = True
        Exit Sub
    End If
    
    lblPremioLiquidoCalc.Caption = "R$ " & Format(Val_Pr_Tot_Orig, "###,###,##0.00")
    'txtValPremioInformado.Text = Format(Val_Pr_Liq, "###,###,##0.00")
    If Flg_Tudo_Premio_Informado = True Then
        txtValPremioInformado.Text = Format(mobjstcPed.VAL_PR, "###,###,##0.00")
        Flg_PremioLiquidoBateComSoma = True
    Else
        txtValPremioInformado.Text = ""
    End If
    
End Sub

Public Sub gpInstanciarForm(pobjstcPed As stcA46V702B)
    Set mobjstcPed = pobjstcPed
    Load Me
    CarregaForm
End Sub

Private Sub ConfigGrid()
    flxCoberturas.Rows = 1
    flxCoberturas.Cols = 6
    flxCoberturas.Enabled = True
    
    flxCoberturas.TextMatrix(0, 0) = ""
    flxCoberturas.ColWidth(0) = 250
    
    flxCoberturas.TextMatrix(0, T46V202_Grid.Item) = "Item"
    flxCoberturas.ColAlignment(T46V202_Grid.Item) = flexAlignCenterCenter
    flxCoberturas.ColWidth(T46V202_Grid.Item) = 470
    
    flxCoberturas.TextMatrix(0, T46V202_Grid.Codigo) = "Cód"
    flxCoberturas.ColAlignment(T46V202_Grid.Codigo) = flexAlignCenterCenter
    flxCoberturas.ColWidth(T46V202_Grid.Codigo) = 470
    
    flxCoberturas.TextMatrix(0, T46V202_Grid.Descricao) = "Cobertura"
    flxCoberturas.ColAlignment(T46V202_Grid.Descricao) = flexAlignCenterCenter
    flxCoberturas.ColWidth(T46V202_Grid.Descricao) = 4750
    
    flxCoberturas.TextMatrix(0, T46V202_Grid.Val_pr_original) = "Prêmio calculado"
    flxCoberturas.ColAlignment(T46V202_Grid.Val_pr_original) = flexAlignCenterCenter
    flxCoberturas.ColWidth(T46V202_Grid.Val_pr_original) = 1375
    
    flxCoberturas.TextMatrix(0, T46V202_Grid.VAL_PR_Informado) = "Prêmio informado"
    flxCoberturas.ColAlignment(T46V202_Grid.VAL_PR_Informado) = flexAlignCenterCenter
    flxCoberturas.ColWidth(T46V202_Grid.VAL_PR_Informado) = 1375
    

End Sub

Private Sub cmdDistribuir_Click()
    ' Varrer as coberturas para definir o novo premio informado.
    Dim pobjstcPedLoc As stcA46V708B
    Dim pobjstcPedCobert As stcA46V704B
    Dim l As Long
    Dim c As Long
    Dim FatorCobert As Double
    Dim Val_Pr_Total_Informado As Double
    Dim Soma_Coberturas As Double
    Dim Diferenca_Soma_Coberturas As Double
    
    If Len(txtValPremioInformado.Text) = 0 Then
        MsgBox "Você precisa digitar um valor antes de distribuir", vbExclamation, "Aviso"
        Exit Sub
    End If
    
    If IsNumeric(txtValPremioInformado.Text) = False Then
        MsgBox "O valor do premio informado precisa ser numérico", vbExclamation, "Aviso"
        Exit Sub
    End If
    
    Val_Pr_Total_Informado = CDbl(txtValPremioInformado.Text)
    
    For l = 1 To mobjstcPed.ColecaoPedLoc.Count
        Set pobjstcPedLoc = mobjstcPed.ColecaoPedLoc(l)
        For c = 1 To pobjstcPedLoc.ColecaoPedCobert.Count
            Set pobjstcPedCobert = pobjstcPedLoc.ColecaoPedCobert(c)
            
            ' Obter o % do premio desta cobertura em relação ao premio total
            If pobjstcPedCobert.PremioInformado = False Then
                FatorCobert = pobjstcPedCobert.VAL_PR / Val_Pr_Tot_Orig
            Else
                FatorCobert = pobjstcPedCobert.Val_Pr_Orig / Val_Pr_Tot_Orig
            End If
            
            ' Aplicar o premio informado
            If pobjstcPedCobert.PremioInformado = False Then
                pobjstcPedCobert.Val_Pr_Orig = pobjstcPedCobert.VAL_PR
                pobjstcPedCobert.PremioInformado = True
            End If
            
            pobjstcPedCobert.VAL_PR = Round(FatorCobert * Val_Pr_Total_Informado, 2)
            Soma_Coberturas = Soma_Coberturas + pobjstcPedCobert.VAL_PR
        Next c
    Next l
    
    mobjstcPed.VAL_PR = Val_Pr_Total_Informado
    mobjstcPed.VAL_PR_Informado = Val_Pr_Total_Informado
    
    ' Ver se a soma é igual ao valor informado. Se não for, ajustar
    Diferenca_Soma_Coberturas = Val_Pr_Total_Informado - Soma_Coberturas
    
    If Diferenca_Soma_Coberturas <> 0# Then
        mobjstcPed.ColecaoPedLoc(1).ColecaoPedCobert(1).VAL_PR = mobjstcPed.ColecaoPedLoc(1).ColecaoPedCobert(1).VAL_PR + Diferenca_Soma_Coberturas
        'MsgBox "Houve uma diferença de R$ " & Format(Diferenca_Soma_Coberturas, "standard") & " no arredondamento. A diferença foi somada à cobertura básica do primeiro item.", vbInformation, "Diferença de arredondamento"
    End If
    
    ' Chamar novamente o PreencheGrid
    CarregaGrid

End Sub

Private Sub cmdLimpar_Click()
    Dim pobjstcPedLoc As stcA46V708B
    Dim pobjstcPedCobert As stcA46V704B
    Dim l As Long
    Dim c As Long

    For l = 1 To mobjstcPed.ColecaoPedLoc.Count
        Set pobjstcPedLoc = mobjstcPed.ColecaoPedLoc(l)
        For c = 1 To pobjstcPedLoc.ColecaoPedCobert.Count
            Set pobjstcPedCobert = pobjstcPedLoc.ColecaoPedCobert(c)
            
            If pobjstcPedCobert.PremioInformado <> False Then
                pobjstcPedCobert.VAL_PR = pobjstcPedCobert.Val_Pr_Orig
                pobjstcPedCobert.Val_Pr_Orig = 0
                pobjstcPedCobert.PremioInformado = False
            End If
            
        Next c
    Next l
    
    txtValPremioInformado.Text = ""
    mobjstcPed.VAL_PR_Informado = 0
    
    ' Chamar novamente o PreencheGrid
    CarregaGrid
End Sub

Private Sub cmdSalvar_Click()
    
    If txtValPremioInformado.Text = "" Then
        mobjstcPed.VAL_PR = Val_Pr_Liq
        mobjstcPed.VAL_PR_Informado = 0
    Else
        If IsNumeric(txtValPremioInformado.Text) = False Then
            MsgBox "O valor do premio informado precisa ser numérico", vbExclamation, "Aviso"
            txtValPremioInformado.Text = ""
            Exit Sub
        End If
        
        If Flg_PremioLiquidoBateComSoma = False Then
            If MsgBox("O valor do prêmio liquido informado não corresponde a soma dos premios das coberturas." & vbCrLf & "Tem certeza que deseja manter assim?", vbOKCancel, "Diferença de prêmio") = vbCancel Then Exit Sub
        End If
        mobjstcPed.VAL_PR = CDbl(txtValPremioInformado.Text)
        mobjstcPed.VAL_PR_Informado = CDbl(txtValPremioInformado.Text)
    End If
    
    Flg_ClicouSalvar = True
    Unload Me
End Sub

Private Sub Form_Load()
    Flg_ClicouSalvar = False
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If Flg_ClicouSalvar = True Then Exit Sub
    
    If txtValPremioInformado.Text = "" Then
        mobjstcPed.VAL_PR = Val_Pr_Liq
        mobjstcPed.VAL_PR_Informado = 0
    Else
        If IsNumeric(txtValPremioInformado.Text) = False Then
            MsgBox "O valor do premio informado precisa ser numérico", vbExclamation, "Aviso"
            txtValPremioInformado.Text = ""
            Exit Sub
        End If
        
        If Flg_PremioLiquidoBateComSoma = False Then
            If MsgBox("O valor do prêmio liquido informado não corresponde a soma dos premios das coberturas." & vbCrLf & "Tem certeza que deseja manter assim?", vbOKCancel, "Diferença de prêmio") = vbCancel Then
                Cancel = True
                Exit Sub
            End If
        End If
        mobjstcPed.VAL_PR = CDbl(txtValPremioInformado.Text)
        mobjstcPed.VAL_PR_Informado = CDbl(txtValPremioInformado.Text)
    End If

End Sub

Private Sub txtValPremioInformado_LostFocus()
    Dim ValPrInf As Double
    
    txtValPremioInformado.Text = Format(txtValPremioInformado.Text, "standard")
    
    If Len(txtValPremioInformado.Text) > 0 Then
        If IsNumeric(txtValPremioInformado.Text) = False Then
            MsgBox "O valor do premio informado precisa ser numérico", vbExclamation, "Aviso"
            txtValPremioInformado.Text = ""
            Exit Sub
        End If
    
        ValPrInf = CDbl(txtValPremioInformado.Text)
    Else
        ValPrInf = 0
    End If
    
    If ValPrInf <> Val_Pr_Liq Then
        Flg_PremioLiquidoBateComSoma = False
    End If
    
End Sub
