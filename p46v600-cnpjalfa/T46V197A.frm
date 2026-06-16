VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmT46V197A 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Bicicletas"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9435
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   9435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtDescricao 
      ForeColor       =   &H00800000&
      Height          =   2055
      Left            =   960
      MaxLength       =   500
      MultiLine       =   -1  'True
      TabIndex        =   3
      ToolTipText     =   "Nome do segurado."
      Top             =   390
      Width           =   8370
   End
   Begin VB.TextBox txtValor 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   960
      MaxLength       =   60
      TabIndex        =   5
      ToolTipText     =   "Nome do segurado."
      Top             =   2490
      Width           =   1320
   End
   Begin VB.TextBox txtItem 
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   960
      MaxLength       =   60
      TabIndex        =   1
      ToolTipText     =   "Nome do segurado."
      Top             =   60
      Width           =   600
   End
   Begin Threed.SSCommand cmdOk 
      Height          =   360
      Left            =   60
      TabIndex        =   6
      ToolTipText     =   "Novo bem."
      Top             =   2970
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   635
      _StockProps     =   78
      Caption         =   "&Ok"
      ForeColor       =   8388608
   End
   Begin Threed.SSCommand cmdCancelar 
      Height          =   360
      Left            =   1110
      TabIndex        =   7
      ToolTipText     =   "Volta à proposta."
      Top             =   2970
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   635
      _StockProps     =   78
      Caption         =   "&Cancelar"
      ForeColor       =   8388608
   End
   Begin VB.Label lblDescricao 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Descrição: "
      ForeColor       =   &H00800000&
      Height          =   2055
      Left            =   60
      TabIndex        =   2
      ToolTipText     =   "Nome do segurado."
      Top             =   390
      Width           =   900
   End
   Begin VB.Label lblValor 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Valor: "
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   60
      TabIndex        =   4
      ToolTipText     =   "Nome do segurado."
      Top             =   2490
      Width           =   900
   End
   Begin VB.Label lblItem 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Item: "
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   60
      TabIndex        =   0
      ToolTipText     =   "Nome do segurado."
      Top             =   60
      Width           =   900
   End
End
Attribute VB_Name = "frmT46V197A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Bicicleta As stcA46V744B

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Call gpHabSalvar
    Bicicleta.NOSSO_NUMERO = Bicicleta.ObjPai.NOSSO_NUMERO
    Bicicleta.Cod_Org = Bicicleta.ObjPai.Cod_Org
    Bicicleta.Cod_Ramo = Bicicleta.ObjPai.Cod_Ramo
    Bicicleta.NUM_PI = Bicicleta.ObjPai.NUM_PI
    Bicicleta.Num_Ped = Bicicleta.ObjPai.Num_Ped
    Bicicleta.Num_Apol = Bicicleta.ObjPai.Num_Apol
    Bicicleta.Num_Endo = Bicicleta.ObjPai.Num_Endo
    Bicicleta.Num_Item = Bicicleta.ObjPai.Num_Item
    
    Bicicleta.Cod_Seq = txtItem.Text
    Bicicleta.Descricao = txtDescricao.Text
    Bicicleta.Valor = gM46V999.CDblx(txtValor.Text)
    Unload Me
End Sub

Public Sub Alterar(ByRef pBicicleta As stcA46V744B)
    Set Bicicleta = pBicicleta
    txtItem.Text = Bicicleta.Cod_Seq
    txtDescricao.Text = Bicicleta.Descricao
    txtValor.Text = Bicicleta.Valor
End Sub

Public Sub Incluir(ByRef pBicicleta As stcA46V744B)
    Set Bicicleta = pBicicleta
    Bicicleta.Cod_Org = Bicicleta.ObjPai.Cod_Org
    Bicicleta.Cod_Seq = Bicicleta.ObjPai.ColecaoPedLocBicicleta.Count + 1
    txtItem.Text = Bicicleta.Cod_Seq
End Sub

Private Sub txtValor_GotFocus()
    With txtValor
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
    KeyAscii = ConsisteKeyNum(KeyAscii, ",")

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtValor_LostFocus()
    Dim i As Integer
    Dim lI As Integer

    With txtValor
        If gM46V999.gfPreenchido(.Text) Then
            If IsNumeric(.Text) Then
                .Text = Format(.Text, "standard")
                If .Text > 2147483647 Then
                    MsgBox "Valor em risco excedeu o limite!", vbInformation, "Atenção"
                    txtValor.SetFocus
                End If
            Else
                .Text = Format(0, "standard")
            End If
        Else
            .Text = Format(0, "standard")
        End If
    End With

End Sub
