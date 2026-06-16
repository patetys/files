VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmT46V153A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro Técnico da Especificação da Proposta - Editor de textos"
   ClientHeight    =   7065
   ClientLeft      =   1035
   ClientTop       =   1005
   ClientWidth     =   10020
   Icon            =   "T46V153A.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   10020
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CMDialog1 
      Left            =   8370
      Top             =   270
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9045
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V153A.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V153A.frx":0426
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V153A.frx":087A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V153A.frx":0996
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V153A.frx":0AB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V153A.frx":0BCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V153A.frx":0CE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V153A.frx":0DF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V153A.frx":0F0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V153A.frx":101E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V153A.frx":1472
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V153A.frx":18C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V153A.frx":19DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V153A.frx":1AEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V153A.frx":209A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V153A.frx":24EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V153A.frx":2ABA
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V153A.frx":300E
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V153A.frx":3328
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V153A.frx":3642
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V153A.frx":3754
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V153A.frx":3866
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Threed.SSPanel PanEditor 
      Height          =   6075
      Left            =   0
      TabIndex        =   6
      Top             =   900
      Width           =   10005
      _Version        =   65536
      _ExtentX        =   17648
      _ExtentY        =   10716
      _StockProps     =   15
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
      BorderWidth     =   2
      BevelOuter      =   1
      BevelInner      =   1
      Enabled         =   0   'False
      Begin RichTextLib.RichTextBox tmpeditor 
         Height          =   5595
         Left            =   405
         TabIndex        =   7
         Top             =   180
         Visible         =   0   'False
         Width           =   9405
         _ExtentX        =   16589
         _ExtentY        =   9869
         _Version        =   393217
         ScrollBars      =   2
         TextRTF         =   $"T46V153A.frx":3978
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
      Begin RichTextLib.RichTextBox editor 
         Height          =   6000
         Left            =   540
         TabIndex        =   8
         Top             =   45
         Width           =   9405
         _ExtentX        =   16589
         _ExtentY        =   10583
         _Version        =   393217
         BackColor       =   16777215
         HideSelection   =   0   'False
         ScrollBars      =   3
         RightMargin     =   9450
         TextRTF         =   $"T46V153A.frx":3A03
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
   Begin Threed.SSPanel PnlTexto 
      Height          =   555
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   10005
      _Version        =   65536
      _ExtentX        =   17648
      _ExtentY        =   979
      _StockProps     =   15
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
      BorderWidth     =   1
      BevelOuter      =   0
      BevelInner      =   1
      Alignment       =   0
      Begin Threed.SSCommand cmdBuscaTxt 
         Height          =   480
         Left            =   8145
         TabIndex        =   5
         Top             =   45
         Width           =   1800
         _Version        =   65536
         _ExtentX        =   3175
         _ExtentY        =   847
         _StockProps     =   78
         Caption         =   "Selecionar texto padrão"
         ForeColor       =   12582912
      End
      Begin VB.Label lblTipTexto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "TEXTO EDITÁVEL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   2340
         TabIndex        =   4
         Top             =   90
         Width           =   3915
      End
   End
   Begin MSComctlLib.Toolbar barFerramenta 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10020
      _ExtentX        =   17674
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   45
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SALVAR"
            Object.ToolTipText     =   "Salvar Texto e Retornar (CTRL+S)"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "BOLD"
            Object.ToolTipText     =   "Negrito"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ITALICS"
            Object.ToolTipText     =   "Itálico"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "UNDER"
            Object.ToolTipText     =   "Sublinhado"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "STRIKE"
            Object.ToolTipText     =   "Tachado"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CLOUR"
            Object.ToolTipText     =   "Cor do Texto"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button29 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button30 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button31 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button32 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button33 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button34 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button35 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CUT"
            Object.ToolTipText     =   "Recortar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button36 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "COPY"
            Object.ToolTipText     =   "Copiar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button37 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PASTE"
            Object.ToolTipText     =   "Colar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button38 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button39 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "BULLET"
            Object.ToolTipText     =   "Marcadores"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button40 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ALESQUERDA"
            Object.ToolTipText     =   "Alinhar à esquerda"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button41 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CENTRALIZAR"
            Object.ToolTipText     =   "Centralizar"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button42 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ALDIREITA"
            Object.ToolTipText     =   "Alinhar à direita"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button43 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   11
            Style           =   3
         EndProperty
         BeginProperty Button44 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button45 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SAIR"
            Object.ToolTipText     =   "Fecha sem Salvar (CTRL+F4)"
            ImageIndex      =   18
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.ComboBox cbosize 
         Height          =   315
         Left            =   4620
         TabIndex        =   2
         Text            =   "cbosize"
         Top             =   0
         Width           =   645
      End
      Begin VB.ComboBox cbofont 
         Height          =   315
         Left            =   2400
         Locked          =   -1  'True
         Sorted          =   -1  'True
         TabIndex        =   1
         Text            =   "cbofont"
         Top             =   0
         Width           =   2085
      End
   End
   Begin RichTextLib.RichTextBox RchFinal 
      Height          =   420
      Left            =   9495
      TabIndex        =   9
      Top             =   2655
      Visible         =   0   'False
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   741
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      TextRTF         =   $"T46V153A.frx":3A85
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Arquivo"
      Begin VB.Menu mnuSave 
         Caption         =   "&Salvar"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnu_Line3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Sair"
         Shortcut        =   ^{F4}
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Editar"
      Begin VB.Menu mnu_Undo 
         Caption         =   "Desfazer"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnu_Line 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "&Recortar"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Colar"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copiar"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuClear 
         Caption         =   "Limpa&r"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnu_Tab1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelect 
         Caption         =   "&Selecionar Tudo"
      End
      Begin VB.Menu mnu_Line1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Wrap 
         Caption         =   "Quebra automática de linha"
      End
   End
   Begin VB.Menu mnuOption 
      Caption         =   "&Formatar"
      Begin VB.Menu mnuBold 
         Caption         =   "&Negrito"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuUnderLine 
         Caption         =   "&Sunlinhado"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuItalics 
         Caption         =   "&Itálico"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuStrike 
         Caption         =   "&Tachado"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnu_clour 
         Caption         =   "&Cor"
         Shortcut        =   ^L
      End
   End
   Begin VB.Menu mnu_Search 
      Caption         =   "&Busca"
      Begin VB.Menu mnu_Find 
         Caption         =   "&Localizar"
      End
   End
End
Attribute VB_Name = "frmT46V153A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mTipDat As E_TipDat
Private Enum E_TipDat
    E_Texto = 0
    E_Image = 1
End Enum

Private Enum E_TipFormat
    E_ManipND = 0
    E_ManipNegrito = 1
    E_ManipItalico = 2
    E_ManipSublinhado = 3
    E_ManipTachado = 4
    E_ManipMarcador = 5
    E_ManipCor = 6
    E_ManipCopy = 7
    E_ManipCut = 8
    E_ManipPaste = 9
    E_ManipLimpar = 10
    E_ManipSelTudo = 11
    E_ManipAlinEsq = 12
    E_ManipAlinDir = 13
    E_ManipAlinCent = 14
End Enum
Dim mFlagSalvar As Boolean
Private mFlagSoLeitura As Boolean
Private mCarga As Boolean
Private mAbrePesquisa As Boolean

Dim mstcCotaTexto As stcA46V724B
Dim AtualEspec As e_Tipo_Espec
Private Sub barFerramenta_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case UCase(Button.Key)
    Case "SALVAR"
        mnuSave_Click
    Case "SAIR"
        mnuExit_Click
    Case "BOLD"
        Call mnuBold_Click
    Case "ITALICS"
        Call mnuItalics_Click
    Case "UNDER"
        Call mnuUnderLine_Click
    Case "STRIKE"
        Call mnuStrike_Click
    Case "CLOUR"
        Call mnu_clour_Click
    Case "CUT"
        Call mnuCut_Click
    Case "COPY"
        Call mnuCopy_Click
    Case "PASTE"
        Call mnuPaste_Click
    Case "BULLET"
        If editor.Visible = True Then
            Call mpManipulaRichText(E_ManipMarcador, editor)
        ElseIf tmpeditor.Visible = True Then
            Call mpManipulaRichText(E_ManipMarcador, tmpeditor)
        End If
    Case "ALESQUERDA"
        If editor.Visible = True Then
            Call mpManipulaRichText(E_ManipAlinEsq, editor)
        ElseIf tmpeditor.Visible = True Then
            Call mpManipulaRichText(E_ManipAlinEsq, tmpeditor)
        End If
    Case "CENTRALIZAR"
        If editor.Visible = True Then
            Call mpManipulaRichText(E_ManipAlinCent, editor)
        ElseIf tmpeditor.Visible = True Then
            Call mpManipulaRichText(E_ManipAlinCent, tmpeditor)
        End If
    Case "ALDIREITA"
        If editor.Visible = True Then
            Call mpManipulaRichText(E_ManipAlinDir, editor)
        ElseIf tmpeditor.Visible = True Then
            Call mpManipulaRichText(E_ManipAlinDir, tmpeditor)
        End If
    End Select

End Sub

Public Sub gpInstanciarForm(ByRef pstcCotaTexto As stcA46V724B)
    Dim lMsg As String
    On Error GoTo Sub_ERRO

    Set mstcCotaTexto = pstcCotaTexto
    mFlagSalvar = False

    Me.Show vbModal

SUB_SAIDA:
    If mFlagSalvar = True Then
        Set pstcCotaTexto = mstcCotaTexto
    End If

    Exit Sub

Sub_ERRO:
    If gM46V999.gfPreenchido(lMsg) = False Then lMsg = "frmT06V053A Pré-Loader erro " & Err & " - " & Error
    Call gM46V999.gpGraLog(2, lMsg)

End Sub


Private Sub cmdBuscaTxt_Click()
    frmT46V144A.TxtAtualEspec = mstcCotaTexto.Tipo_Espec
    RchFinal.TextRTF = ""
    frmT46V144A.Show 1
    If RchFinal.Text <> "" Then
        If editor.Visible Then
            editor.SelText = vbCrLf
            editor.SelRTF = RchFinal.TextRTF
        ElseIf tmpeditor.Visible Then
            tmpeditor.SelText = vbCrLf
            tmpeditor.SelRTF = RchFinal.TextRTF
        End If
    End If


End Sub


Private Sub Form_Activate()
    If editor.Visible And editor.Enabled Then
        editor.SetFocus
        editor_KeyUp 0, 0
    ElseIf tmpeditor.Visible And tmpeditor.Enabled Then
        tmpeditor.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Dim lConta As Integer
    Dim lMensagem As String

    Call gM46V999.gpCenForm(Me)
    mCarga = True
    mAbrePesquisa = False

    editor.SelStart = 1
    editor.SelLength = Len(editor.Text)
    editor.SelFontName = gM46V111.gFonteEspecificacao
    editor.SelFontSize = 9
    editor.SelStart = 1
    editor.SelLength = 0

    tmpeditor.SelStart = 1
    tmpeditor.SelLength = Len(tmpeditor.Text)
    tmpeditor.SelFontName = gM46V111.gFonteEspecificacao
    tmpeditor.SelFontSize = 9
    tmpeditor.SelStart = 1
    tmpeditor.SelLength = 0


    'Carrega as informações na tela

    lblTipTexto.Caption = mstcCotaTexto.Nom_Espec
    mFlagSoLeitura = False
    PanEditor.Enabled = True

    If Trim(mstcCotaTexto.Dsc_Espec) <> "" Then
        editor.TextRTF = mstcCotaTexto.Dsc_Espec
        tmpeditor.TextRTF = editor.TextRTF
    End If

    'Define a largura e altura dos controles RichText
    PanEditor.Left = 0
    PnlTexto.Left = 0
    lblTipTexto.Width = PnlTexto.Width - 90
    lblTipTexto.Left = 40

    PanEditor.Width = Me.Width - 50
    editor.Left = 0
    tmpeditor.Left = 0
    editor.Top = 45
    tmpeditor.Top = 45

    tmpeditor.Width = PanEditor.Width - 20
    editor.Width = PanEditor.Width - 20

    tmpeditor.Height = PanEditor.Height - 10  'Me.Height
    editor.Height = PanEditor.Height - 10      'Me.Height

    'Carrega os combos de Font e Size
    For lConta = 8 To 20 Step 1
        cbosize.AddItem (lConta)
    Next lConta
    cbosize.ListIndex = 1

    For lConta = 0 To Printer.FontCount - 1    ' Numero de fontes disponíveis
        cbofont.AddItem Printer.Fonts(lConta)  ' Inclui cada fonte no combo
    Next lConta
    cbofont.ListIndex = 1

    Call gM46V999.gpCenForm(Me)

    mCarga = False
End Sub

Private Sub Form_Resize()
'    editor.Width = Me.Width - 175
'   tmpeditor.Width = Me.Width - 175

    PanEditor.Width = Me.Width - 50
    tmpeditor.Width = PanEditor.Width - 20
    tmpeditor.Height = PanEditor.Height - 10  'Me.Height
    editor.Width = PanEditor.Width - 20
    editor.Height = PanEditor.Height - 10      'Me.Height


End Sub

Private Sub cbofont_Click()
    If editor.Visible = True Then
        editor.SelFontName = cbofont.Text
    ElseIf tmpeditor.Visible = True Then
        tmpeditor.SelFontName = cbofont.Text
    End If
End Sub

Private Sub cbofont_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cbosize_Click()
    If editor.Visible = True Then
        editor.SelFontSize = Val(cbosize.Text)
    ElseIf tmpeditor.Visible = True Then
        tmpeditor.SelFontSize = Val(cbosize.Text)
    End If
End Sub

Private Sub cbosize_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub editor_KeyDown(KeyCode As Integer, Shift As Integer)
    If mTipDat = E_Image Then
        mTipDat = E_Texto
        Exit Sub
    End If
    If KeyCode = 117 Then
        editor.SelText = UCase$(editor.SelText)
    ElseIf KeyCode = 118 Then
        editor.SelText = LCase$(editor.SelText)
    End If
End Sub

Private Sub editor_KeyPress(KeyAscii As Integer)
    If mTipDat = E_Image Then
        mTipDat = E_Texto
        Exit Sub
    End If
    ' cbosize.Text = editor.SelFontSize
    ' cbofont.Text = editor.SelFontName
    editor.SelFontSize = cbosize.Text
    editor.SelFontName = cbofont.Text
End Sub
Private Sub editor_KeyUp(KeyCode As Integer, Shift As Integer)
    If mTipDat = E_Image Then
        mTipDat = E_Texto
        Exit Sub
    End If

    If gM46V999.gfPreenchido(editor.SelFontName) = True Then cbofont.Text = editor.SelFontName
    If gM46V999.gfPreenchido(editor.SelFontSize) = True Then cbosize.Text = editor.SelFontSize

End Sub

Private Sub editor_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If mTipDat = E_Image Then
        mTipDat = E_Texto
        Exit Sub
    End If
    editor_KeyUp 0, 0
End Sub

Private Sub mpManipulaRichText(ByVal pTipFormat As E_TipFormat, _
                               ByRef pObjTexto As RichTextBox)
    Dim lMsgErr As String

    Err = 0
    On Error GoTo Sub_ERRO

    Select Case pTipFormat
    Case E_ManipND
        Exit Sub
    Case E_ManipNegrito
        If pObjTexto.SelBold = True Then
            pObjTexto.SelBold = False
        Else
            pObjTexto.SelBold = True
        End If
    Case E_ManipItalico
        If pObjTexto.SelItalic = True Then
            pObjTexto.SelItalic = False
        Else
            pObjTexto.SelItalic = True
        End If
    Case E_ManipSublinhado
        If pObjTexto.SelUnderline = True Then
            pObjTexto.SelUnderline = False
        Else
            pObjTexto.SelUnderline = True
        End If
    Case E_ManipTachado
        If pObjTexto.SelStrikeThru = True Then
            pObjTexto.SelStrikeThru = False
        Else
            pObjTexto.SelStrikeThru = True
        End If
    Case E_ManipMarcador
        If pObjTexto.SelBullet = True Then
            pObjTexto.SelBullet = False
        Else
            pObjTexto.SelBullet = True
        End If
    Case E_ManipCor
        CMDialog1.ShowColor
        pObjTexto.SelColor = CMDialog1.color
    Case E_ManipCopy
        Clipboard.SetText pObjTexto.SelText
    Case E_ManipCut
        Clipboard.SetText pObjTexto.SelText
        pObjTexto.SelText = ""
    Case E_ManipPaste
        pObjTexto.SelText = Clipboard.GetText()
    Case E_ManipLimpar
        pObjTexto.SelText = ""
    Case E_ManipSelTudo
        pObjTexto.SelStart = 0
        pObjTexto.SelLength = Len(pObjTexto.Text)
    Case E_ManipAlinEsq
        pObjTexto.SelAlignment = rtfLeft
    Case E_ManipAlinDir
        pObjTexto.SelAlignment = rtfRight
    Case E_ManipAlinCent
        pObjTexto.SelAlignment = rtfCenter
    Case Else
        lMsgErr = "mpManipulaRichText ERRO - Tipo de formatação desconhecido (" & pTipFormat & ")."
        GoTo Sub_ERRO
    End Select

SUB_SAIDA:
    Err = 0
    On Error GoTo 0
    Exit Sub

Sub_ERRO:
    If gM46V999.gfPreenchido(lMsgErr) = False Then lMsgErr = "mpManipulaRichText ERRO " & Err & " - " & Error
    Call gM46V999.gpGraLog(2, lMsgErr)
    GoTo SUB_SAIDA

End Sub



Private Sub mnu_clour_Click()
    If editor.Visible = True Then
        Call mpManipulaRichText(E_ManipCor, editor)
    ElseIf tmpeditor.Visible = True Then
        Call mpManipulaRichText(E_ManipCor, tmpeditor)
    End If
End Sub

Private Sub mnu_Find_Click()
    Dim FoundPos As Integer
    Dim FoundLine As Integer
    Dim serstring As String
    ' Procura o texto specificado no InputBox.

    serstring = InputBox("Digite o texto a ser procurado :", "Procurar expressão", serstring)
    If editor.Visible = True Then
        FoundPos = editor.Find(serstring, , , rtfWholeWord)
        If FoundPos <> -1 Then
            ' Retorna o numero da linha contendo o texto localizado.
            FoundLine = editor.GetLineFromChar(FoundPos)
            MsgBox "Expressão encontrada na linha " & CStr(FoundLine)
        Else
            MsgBox "Expressão não encontrada."
        End If
    ElseIf tmpeditor.Visible = True Then
        FoundPos = tmpeditor.Find(serstring, , , rtfWholeWord)
        If FoundPos <> -1 Then
            ' Retorna o numero da linha contendo o texto localizado.
            FoundLine = tmpeditor.GetLineFromChar(FoundPos)
            MsgBox "Expressão encontrada na linha " & CStr(FoundLine)
        Else
            MsgBox "Expressão não encontrada."
        End If
    End If

End Sub

Private Sub mnu_undo_Click()
    If editor.Visible = True Then
        Call mpManipulaRichText(E_ManipPaste, editor)
    ElseIf tmpeditor.Visible = True Then
        Call mpManipulaRichText(E_ManipPaste, tmpeditor)
    End If
End Sub

Private Sub mnu_Wrap_Click()
'Inverte o Obtejo RichText ativo
    If editor.Visible = True Then
        tmpeditor.TextRTF = editor.TextRTF
        mnu_Wrap.Checked = True
        tmpeditor.Visible = True
        editor.Visible = False
    ElseIf tmpeditor.Visible = True Then
        editor.TextRTF = tmpeditor.TextRTF
        mnu_Wrap.Checked = False
        editor.Visible = True
        tmpeditor.Visible = False
    End If
End Sub

Private Sub mnuBold_Click()
    If editor.Visible = True Then
        Call mpManipulaRichText(E_ManipNegrito, editor)
    ElseIf tmpeditor.Visible = True Then
        Call mpManipulaRichText(E_ManipNegrito, tmpeditor)
    End If
End Sub

Private Sub mnuClear_Click()
    If editor.Visible = True Then
        Call mpManipulaRichText(E_ManipLimpar, editor)
    ElseIf tmpeditor.Visible = True Then
        Call mpManipulaRichText(E_ManipLimpar, tmpeditor)
    End If
End Sub

Private Sub mnuCopy_Click()
    If editor.Visible = True Then
        Call mpManipulaRichText(E_ManipCopy, editor)
    ElseIf tmpeditor.Visible = True Then
        Call mpManipulaRichText(E_ManipCopy, tmpeditor)
    End If
End Sub

Private Sub mnuCut_Click()
    If editor.Visible = True Then
        Call mpManipulaRichText(E_ManipCut, editor)
    ElseIf tmpeditor.Visible = True Then
        Call mpManipulaRichText(E_ManipCut, tmpeditor)
    End If
End Sub

Private Sub mnuExit_Click()
    Call mpCancelar
    Unload Me
End Sub

Private Sub mnuItalics_Click()
    If editor.Visible = True Then
        Call mpManipulaRichText(E_ManipItalico, editor)
    ElseIf tmpeditor.Visible = True Then
        Call mpManipulaRichText(E_ManipItalico, tmpeditor)
    End If
End Sub

Private Sub mnuPaste_Click()
    If editor.Visible = True Then
        Call mpManipulaRichText(E_ManipPaste, editor)
    ElseIf tmpeditor.Visible = True Then
        Call mpManipulaRichText(E_ManipPaste, tmpeditor)
    End If
End Sub




Private Sub mnuSalvarVAtual_Click()

End Sub


Private Sub mnuSalvarVNova_Click()

End Sub


Private Sub mnuSave_Click()
    Call mpSalvar
    Unload Me
End Sub

Private Sub mnuSelect_Click()
    If editor.Visible = True Then
        Call mpManipulaRichText(E_ManipSelTudo, editor)
    ElseIf tmpeditor.Visible = True Then
        Call mpManipulaRichText(E_ManipSelTudo, tmpeditor)
    End If
End Sub

Private Sub mnuStrike_Click()
    If editor.Visible = True Then
        Call mpManipulaRichText(E_ManipTachado, editor)
    ElseIf tmpeditor.Visible = True Then
        Call mpManipulaRichText(E_ManipTachado, tmpeditor)
    End If
End Sub

Private Sub mnuUnderLine_Click()
    If editor.Visible = True Then
        Call mpManipulaRichText(E_ManipSublinhado, editor)
    ElseIf tmpeditor.Visible = True Then
        Call mpManipulaRichText(E_ManipSublinhado, tmpeditor)
    End If
End Sub

Private Sub mpSalvar(Optional pSalvar As Boolean = True)
    Dim lstcTexto As stcA46V724B
    Dim lMsg As String

    Err = 0
    On Error Resume Next

    mFlagSalvar = False
    Set lstcTexto = New stcA46V724B
    lstcTexto.Tipo_Espec = mstcCotaTexto.Tipo_Espec
    lstcTexto.Nom_Espec = mstcCotaTexto.Nom_Espec
    lstcTexto.Dsc_Espec = mstcCotaTexto.Dsc_Espec

    If editor.Visible = True Then
        editor.SelStart = 1
        editor.SelLength = Len(editor.Text)
        editor.SelFontName = gM46V111.gFonteEspecificacao
        editor.SelFontSize = 9
        DoEvents
        If Trim(editor.Text) = "" Then
            lstcTexto.Dsc_Espec = ""
        Else
            lstcTexto.Dsc_Espec = editor.TextRTF
        End If
    ElseIf tmpeditor.Visible = True Then
        tmpeditor.SelStart = 1
        tmpeditor.SelLength = Len(tmpeditor.Text)
        tmpeditor.SelFontName = gM46V111.gFonteEspecificacao
        tmpeditor.SelFontSize = 9
        DoEvents
        If Trim(tmpeditor.Text) = "" Then
            lstcTexto.Dsc_Espec = ""
        Else
            lstcTexto.Dsc_Espec = tmpeditor.TextRTF
        End If
    Else
        Call gM46V999.gpGraLog(2, "Não foi possivel salvar as o texto editado")
        GoTo SUB_SAIDA
    End If

    mFlagSalvar = pSalvar
    If pSalvar = True Then
        Set mstcCotaTexto = lstcTexto
        Call gpLimPremios
    Else
        GoTo SUB_SAIDA
    End If

SUB_SAIDA:
    Err = 0
    On Error GoTo 0
    Exit Sub

Sub_ERRO:
    MsgBox Err & "-" & Error
    GoTo SUB_SAIDA
End Sub

Private Sub mpCancelar()
    mFlagSalvar = False
End Sub

