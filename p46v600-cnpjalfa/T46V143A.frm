VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmT46V143A 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Especificação do Risco"
   ClientHeight    =   7470
   ClientLeft      =   1560
   ClientTop       =   885
   ClientWidth     =   8955
   Icon            =   "T46V143A.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   Begin RichTextLib.RichTextBox RchFinal 
      Height          =   420
      Left            =   5220
      TabIndex        =   15
      Top             =   45
      Visible         =   0   'False
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   741
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      TextRTF         =   $"T46V143A.frx":030A
   End
   Begin MSComctlLib.ImageList imgLista 
      Left            =   7335
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V143A.frx":038C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V143A.frx":06A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V143A.frx":09C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V143A.frx":0CDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V143A.frx":0FF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V143A.frx":130E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V143A.frx":1920
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V143A.frx":1C3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V143A.frx":2264
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V143A.frx":2876
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V143A.frx":2B90
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V143A.frx":2EAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "T46V143A.frx":31C4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   6975
      LargeChange     =   500
      Left            =   8730
      SmallChange     =   2000
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   450
      Width           =   225
   End
   Begin MSComctlLib.Toolbar barFerramentas 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   847
      ButtonWidth     =   714
      ButtonHeight    =   688
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "imgLista"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Negrito"
            Object.ToolTipText     =   "Negrito"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Italico"
            Object.ToolTipText     =   "Itálico"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Sublinhado"
            Object.ToolTipText     =   "Sublinhado"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1000
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Padrao"
            Object.ToolTipText     =   "Busca Padrão"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1000
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Visualizar"
            Object.ToolTipText     =   "Visualizar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Remover"
            Object.ToolTipText     =   "Remover Especificação"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1500
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Diminuir"
            Object.ToolTipText     =   "Diminuir"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Maximizar"
            Object.ToolTipText     =   "Maximizar"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1000
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin RichTextLib.RichTextBox RchTemp 
         Height          =   420
         Left            =   3105
         TabIndex        =   35
         Top             =   0
         Visible         =   0   'False
         Width           =   510
         _ExtentX        =   900
         _ExtentY        =   741
         _Version        =   393217
         Enabled         =   -1  'True
         HideSelection   =   0   'False
         ScrollBars      =   3
         TextRTF         =   $"T46V143A.frx":36C3
      End
   End
   Begin VB.PictureBox PicPanel 
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   9500
      Left            =   0
      ScaleHeight     =   9495
      ScaleWidth      =   8745
      TabIndex        =   17
      Top             =   400
      Width           =   8750
      Begin VB.Frame FrmObservacao 
         Caption         =   "Observações"
         ForeColor       =   &H00004000&
         Height          =   735
         Left            =   0
         TabIndex        =   33
         Top             =   8760
         Width           =   8700
         Begin RichTextLib.RichTextBox RchObservacao 
            Height          =   510
            Left            =   45
            TabIndex        =   14
            Top             =   180
            Width           =   8600
            _ExtentX        =   15161
            _ExtentY        =   900
            _Version        =   393217
            Enabled         =   -1  'True
            HideSelection   =   0   'False
            ScrollBars      =   3
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"T46V143A.frx":3745
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame FrmOutrasInf 
         Caption         =   "Outras Informações"
         ForeColor       =   &H00800000&
         Height          =   735
         Left            =   0
         TabIndex        =   32
         Top             =   8025
         Width           =   8700
         Begin RichTextLib.RichTextBox RchOutrasInf 
            Height          =   510
            Left            =   45
            TabIndex        =   13
            Top             =   180
            Width           =   8600
            _ExtentX        =   15161
            _ExtentY        =   900
            _Version        =   393217
            Enabled         =   -1  'True
            HideSelection   =   0   'False
            ScrollBars      =   3
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"T46V143A.frx":37C7
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame FrmCondicao 
         Caption         =   "Condições Especiais"
         ForeColor       =   &H00004000&
         Height          =   735
         Left            =   0
         TabIndex        =   31
         Top             =   7290
         Width           =   8700
         Begin RichTextLib.RichTextBox RchCondicao 
            Height          =   510
            Left            =   45
            TabIndex        =   12
            Top             =   180
            Width           =   8600
            _ExtentX        =   15161
            _ExtentY        =   900
            _Version        =   393217
            Enabled         =   -1  'True
            HideSelection   =   0   'False
            ScrollBars      =   3
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"T46V143A.frx":3849
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame FrmDados 
         Height          =   1320
         Left            =   0
         TabIndex        =   27
         Top             =   90
         Width           =   8700
         Begin VB.TextBox TxtLocalExposicao 
            Height          =   285
            Left            =   90
            MaxLength       =   80
            TabIndex        =   3
            Top             =   900
            Width           =   8475
         End
         Begin VB.TextBox TxtPrazoManut 
            Height          =   285
            Left            =   1620
            MaxLength       =   10
            TabIndex        =   1
            Top             =   360
            Width           =   1230
         End
         Begin VB.TextBox TxtPrazoObra 
            Height          =   285
            Left            =   90
            MaxLength       =   10
            TabIndex        =   0
            Top             =   360
            Width           =   1275
         End
         Begin VB.TextBox TxtContratante 
            Height          =   285
            Left            =   3105
            MaxLength       =   50
            TabIndex        =   2
            Top             =   360
            Width           =   5460
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Local de Exposição"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   90
            TabIndex        =   34
            Top             =   675
            Width           =   1395
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Contratante"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   3105
            TabIndex        =   30
            Top             =   135
            Width           =   825
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Prazo Manutenção"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   1575
            TabIndex        =   29
            Top             =   135
            Width           =   1350
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Prazo da Obra"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   28
            Top             =   135
            Width           =   1020
         End
      End
      Begin VB.Frame FrmFranquia 
         Caption         =   "Franquias"
         ForeColor       =   &H00004000&
         Height          =   735
         Left            =   0
         TabIndex        =   26
         Top             =   5820
         Width           =   8700
         Begin RichTextLib.RichTextBox RchFranquia 
            Height          =   510
            Left            =   45
            TabIndex        =   10
            Top             =   180
            Width           =   8600
            _ExtentX        =   15161
            _ExtentY        =   900
            _Version        =   393217
            Enabled         =   -1  'True
            HideSelection   =   0   'False
            ScrollBars      =   3
            RightMargin     =   5
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"T46V143A.frx":38CB
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame FrmCoberturas 
         Caption         =   "Coberturas"
         ForeColor       =   &H00800000&
         Height          =   735
         Left            =   0
         TabIndex        =   25
         Top             =   5085
         Width           =   8700
         Begin RichTextLib.RichTextBox RchCoberturas 
            Height          =   510
            Left            =   45
            TabIndex        =   9
            Top             =   180
            Width           =   8600
            _ExtentX        =   15161
            _ExtentY        =   900
            _Version        =   393217
            Enabled         =   -1  'True
            HideSelection   =   0   'False
            ScrollBars      =   3
            RightMargin     =   5
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"T46V143A.frx":394D
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame FrmEspecificacao 
         Caption         =   "Especificação do Serviço"
         ForeColor       =   &H00004000&
         Height          =   735
         Left            =   0
         TabIndex        =   24
         Top             =   4350
         Width           =   8700
         Begin RichTextLib.RichTextBox RchEspecificacao 
            Height          =   510
            Left            =   45
            TabIndex        =   8
            Top             =   180
            Width           =   8600
            _ExtentX        =   15161
            _ExtentY        =   900
            _Version        =   393217
            Enabled         =   -1  'True
            HideSelection   =   0   'False
            ScrollBars      =   3
            RightMargin     =   5
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"T46V143A.frx":39CF
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame FrmPerimetro 
         Caption         =   "Perímetro da Cobertura"
         ForeColor       =   &H00800000&
         Height          =   705
         Left            =   0
         TabIndex        =   23
         Top             =   3645
         Width           =   8700
         Begin RichTextLib.RichTextBox RchPerimetro 
            Height          =   465
            Left            =   45
            TabIndex        =   7
            Top             =   180
            Width           =   8600
            _ExtentX        =   15161
            _ExtentY        =   820
            _Version        =   393217
            Enabled         =   -1  'True
            HideSelection   =   0   'False
            ScrollBars      =   3
            RightMargin     =   5
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"T46V143A.frx":3A51
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame FrmRCobertos 
         Caption         =   "Riscos Cobertos"
         ForeColor       =   &H00800000&
         Height          =   735
         Left            =   0
         TabIndex        =   21
         Top             =   2175
         Width           =   8700
         Begin RichTextLib.RichTextBox RchRCobertos 
            Height          =   510
            Left            =   45
            TabIndex        =   5
            Top             =   180
            Width           =   8600
            _ExtentX        =   15161
            _ExtentY        =   900
            _Version        =   393217
            Enabled         =   -1  'True
            HideSelection   =   0   'False
            ScrollBars      =   3
            RightMargin     =   5
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"T46V143A.frx":3AD3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame FrmClausula 
         Caption         =   "Cláusulas"
         ForeColor       =   &H00800000&
         Height          =   735
         Left            =   0
         TabIndex        =   20
         Top             =   6555
         Width           =   8700
         Begin RichTextLib.RichTextBox RchClausulas 
            Height          =   510
            Left            =   45
            TabIndex        =   11
            Top             =   180
            Width           =   8600
            _ExtentX        =   15161
            _ExtentY        =   900
            _Version        =   393217
            Enabled         =   -1  'True
            HideSelection   =   0   'False
            ScrollBars      =   3
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"T46V143A.frx":3B55
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame FrmObjeto 
         Caption         =   "Objeto Segurado"
         ForeColor       =   &H00004000&
         Height          =   735
         Left            =   0
         TabIndex        =   19
         Top             =   1440
         Width           =   8700
         Begin RichTextLib.RichTextBox RchObjetos 
            Height          =   510
            Left            =   45
            TabIndex        =   4
            Top             =   180
            Width           =   8600
            _ExtentX        =   15161
            _ExtentY        =   900
            _Version        =   393217
            Enabled         =   -1  'True
            HideSelection   =   0   'False
            ScrollBars      =   3
            RightMargin     =   5
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"T46V143A.frx":3BD7
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame FrmCaracteristica 
         Caption         =   "Característica"
         ForeColor       =   &H00004000&
         Height          =   735
         Left            =   0
         TabIndex        =   18
         Top             =   2910
         Width           =   8700
         Begin RichTextLib.RichTextBox RchCaracteristica 
            Height          =   510
            Left            =   45
            TabIndex        =   6
            Top             =   180
            Width           =   8600
            _ExtentX        =   15161
            _ExtentY        =   900
            _Version        =   393217
            Enabled         =   -1  'True
            HideSelection   =   0   'False
            ScrollBars      =   3
            RightMargin     =   5
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"T46V143A.frx":3C59
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
End
Attribute VB_Name = "frmT46V143A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Dim AtualEspec As e_Tipo_Espec
Dim objWord     As Word.Application

Sub CopiarTexto()
'No evento desejado:
Dim Word As Object
Set Word = CreateObject("Word.Application")

'Abre o documento
Word.Documents.Open "C:\Meus Documentos\" & "NomeDoDocumento.DOC"

'Seleciona o documento inteiro
Word.Documents("NomeDoDocumento.DOC").Select

'Joga o conteúdo da seleção numa caixa de Texto.
'Text1.Text = Word.Selection.Text

End Sub

Sub mpDeletarTemp(ByVal Diretorio As String)
On Error Resume Next
Dim NomeArq As String

    objWord.ActiveDocument.Save
    objWord.ActiveDocument.Close
    objWord.Quit
    Set objWord = Nothing

    NomeArq = Dir(Diretorio & "#Tmp*.ytf", vbArchive)
    Do While NomeArq <> ""   ' Inicia o loop.
        If NomeArq <> "." And NomeArq <> ".." Then
            Kill Diretorio & NomeArq
       End If
       NomeArq = Dir   ' Obtém a próxima entrada.
    Loop
    Err = 0
End Sub

Sub mpJuntarFormatos(ByVal NomeLabel, ByVal ValorRich As RichTextBox)

    If Trim(ValorRich.Text) = "" Then Exit Sub

    RchFinal.SelStart = Len(RchFinal.Text)
    RchFinal.SelBold = True
    RchFinal.SelFontSize = 10
    RchFinal.SelColor = &H0&
    RchFinal.SelFontName = "MS Sans Serif"
    RchFinal.SelText = vbCrLf
    RchFinal.SelText = NomeLabel
    RchFinal.SelText = vbCrLf
    RchFinal.SelBold = False
    RchFinal.SelRTF = ValorRich.TextRTF

End Sub

Sub mpJuntarTextoRTF(ByVal NomeLabel, ByVal ValorRich As String)
    Dim CampoRTF As RichTextBox
    Set CampoRTF = Me.RchTemp
    
    CampoRTF.TextRTF = ValorRich
    
    If Trim(CampoRTF.Text) = "" Then Exit Sub

    RchFinal.SelStart = Len(RchFinal.Text)
    RchFinal.SelBold = True
    RchFinal.SelFontSize = 10
    RchFinal.SelColor = &H0&
    RchFinal.SelFontName = "MS Sans Serif"
    RchFinal.SelText = vbCrLf
    RchFinal.SelText = NomeLabel
    RchFinal.SelText = vbCrLf
    RchFinal.SelBold = False
    RchFinal.SelRTF = CampoRTF.TextRTF

End Sub

Sub mpJuntarTextos(ByVal NomeLabel, ByVal Texto As String, ByVal Quebra As Integer, ByVal Tabs As Integer)
Dim i As Integer
    RchFinal.SelStart = Len(RchFinal.Text)
    RchFinal.SelBold = True
    RchFinal.SelFontSize = 10
    RchFinal.SelColor = &H0&
    RchFinal.SelFontName = "MS Sans Serif"
    RchFinal.SelText = NomeLabel
    RchFinal.SelBold = False
    If Quebra > 0 Then
        For i = 1 To Quebra
            RchFinal.SelText = vbCrLf
        Next
    End If
    If Tabs > 0 Then
        For i = 1 To Tabs
            RchFinal.SelText = vbTab
        Next
    End If
    RchFinal.SelFontSize = 10
    RchFinal.SelText = Texto
    RchFinal.SelText = vbCrLf

End Sub

Sub SubstituirTexto()
    Dim appWord     As Word.Application
    Dim wrdDoc      As Word.Document
    Set appWord = New Word.Application
    Set wrdDoc = appWord.Documents.Open("c:\teste.doc")

    With wrdDoc.ActiveWindow.Selection.Find
        .Execute Findtext:="Robert", replacewith:="FlávioNihi"
    End With

    appWord.Visible = True
    Set appWord = Nothing
    Set wrdDoc = Nothing
End Sub

Sub FindTexto()
'Alterar um texto tanto no principal como no cabecalho e no rodape
    Dim appWord     As Word.Application
    Dim wrdDoc      As Word.Document
    Dim blnTérmino  As Boolean

    Set appWord = New Word.Application

    Set wrdDoc = appWord.Documents.Open("c:\teste.doc")
Dim i
    For i = 1 To 3
        Select Case i
             Case 1
                 wrdDoc.ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
             Case 2
                 wrdDoc.ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
             Case 3
                 wrdDoc.ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
        End Select

        With wrdDoc.ActiveWindow.Selection.Find
             .Text = "Robert"
             .Replacement.Text = "FlávioNihi"
             .Forward = True
             .Wrap = wdFindContinue
             .Execute Replace:=wdReplaceAll
        End With
    Next

    appWord.Visible = True
    Set appWord = Nothing
    Set wrdDoc = Nothing

'----------------
'colocar cabecalho

With ActiveDocument.Sections(1)
    .Headers(wdHeaderFooterPrimary).Range.Text = "Header that I want to set"

    With .Footers(wdHeaderFooterPrimary).Range
        .Text = "Footer Left text" & vbTab & "Footer center" & vbTab
        .Fields.Add Range:=ActiveDocument.Range(.End, .End), Type:=wdFieldPage
        '....
    End With
    
End With
 
 
Dim pDoc As Word.Document
'Set pDoc = WordApp.Documents.Open(FileName)
With pDoc
    .Headers(wdHeaderFooterPrimary).Range.Text = "Header that I want to set"
    With .Footers(wdHeaderFooterPrimary).Range
        .Text = "Footer Left text" & vbTab & "Footer center" & vbTab
        .Fields.Add Range:=ActiveDocument.Range(.End, .End), Type:=wdFieldPage
        '....
    End With

'----------test 2


'Sub SetHeadersFooters()

Dim oApp As Word.Application
Dim oSec As Word.Section
Dim oDoc As Word.Document

'Create a new document in Word
Set oApp = New Word.Application
Set oDoc = oApp.Documents.Add

With oDoc

'=== SECTION 1 ==================================================

'Add two pages to the first section where the first page in the
'section has different headers and footers than the second page
Set oSec = .Sections(1)
oSec.PageSetup.DifferentFirstPageHeaderFooter = True
oSec.Range.InsertAfter "Text on Page 1 (Section 1)"
.Range(oSec.Range.End - 1).InsertBreak wdPageBreak
oSec.Range.InsertAfter "Text on Page 2 (Section 1)"
Debug.Print oSec.Range.ComputeStatistics(wdStatisticCharacters)
MsgBox oSec.Range.ComputeStatistics(wdStatisticCharacters)
'Add the headers/footers for the first section (that contains two
'pages)
oSec.Headers(wdHeaderFooterFirstPage).Range.Text = _
"Page1 -- Section 1 First Page Header"
oSec.Headers(wdHeaderFooterPrimary).Range.Text = _
"Page2 -- Section 1 Primary Header"
oSec.Footers(wdHeaderFooterFirstPage).Range.Text = _
"Page1 -- Section 1 First Page Footer"
oSec.Footers(wdHeaderFooterPrimary).Range.Text = _
"Page2 -- Section 1 Primary Footer"

'=== SECTION 2 ==================================================

'Add a new section containing three pages where the first page in
'the section has different headers and footers than the other two
'pages
.Range(oSec.Range.End - 1).InsertBreak wdSectionBreakNextPage
Set oSec = .Sections(2)
oSec.PageSetup.DifferentFirstPageHeaderFooter = True
oSec.Range.InsertAfter "Text on Page 3 (Section 2)"
.Range(oSec.Range.End - 1).InsertBreak wdPageBreak
oSec.Range.InsertAfter "Text on Page 4 (Section 2)"
.Range(oSec.Range.End - 1).InsertBreak wdPageBreak
oSec.Range.InsertAfter "Text on Page 5 (Section 2)"

'Add the headers/footers for the second section (that contains
'three pages) -- notice that the second and third pages in this
'section will contain the primary header/footer
oSec.Headers(wdHeaderFooterFirstPage).LinkToPrevious = False
oSec.Headers(wdHeaderFooterFirstPage).Range.Text = _
"Page3 -- Section 2 First Page Header"
oSec.Headers(wdHeaderFooterPrimary).LinkToPrevious = False
oSec.Headers(wdHeaderFooterPrimary).Range.Text = _
"Page4and5 -- Section 2 Primary Header"
oSec.Footers(wdHeaderFooterFirstPage).LinkToPrevious = False
oSec.Footers(wdHeaderFooterFirstPage).Range.Text = _
"Page3 -- Section 2 First Page Footer"
oSec.Footers(wdHeaderFooterPrimary).LinkToPrevious = False
oSec.Footers(wdHeaderFooterPrimary).Range.Text = _
"Page4and5 -- Section 2 Primary Footer"

'=== SECTION 3 ==================================================

'Add a new section containing two pages that all have the same
'header/footer
.Range(oSec.Range.End - 1).InsertBreak wdSectionBreakNextPage
Set oSec = .Sections(3)
oSec.PageSetup.DifferentFirstPageHeaderFooter = False
oSec.Range.InsertAfter "Text on Page 6 (Section 3)"
.Range(oSec.Range.End - 1).InsertBreak wdPageBreak
oSec.Range.InsertAfter "Text on Page 7 (Section 3)"

'Add the headers/footers for the third section (that contains
' two pages)
oSec.Headers(wdHeaderFooterPrimary).LinkToPrevious = False
oSec.Headers(wdHeaderFooterPrimary).Range.Text = _
"Page6and7 -- Section 3 Primary Header Only"
oSec.Footers(wdHeaderFooterPrimary).LinkToPrevious = False
oSec.Footers(wdHeaderFooterPrimary).Range.Text = _
"Page6and7 -- Section 3 Primary Footer Only"

'Save the document
.SaveAs App.Path & "\mydoc.doc"

End With

'Make Word visible to examine the document
oApp.Visible = True

End With

'    'With objWord
    '    .Documents.Open NomedoArquivo, AddToRecenteFiles:=False
    '    .ActiveDocument.Paragraphs.Last.Range.bold = False
    '    .ActiveDocument.Paragraphs.Last.Range.Font.Size = 20
    
    '.ActiveDocument.Paragraphs.Last.Range.Font.Name = "??"
    
    '.ActiveDocument.Paragraphs.Last.Range.Font.ColorIndex = 4
    
    '.ActiveDocument.Paragraphs.Last.Range.Text = RI3.TextRTF
    
    'End With
    
   ' Clipboard.SetText Ri.SelText
        'objWord.Selection.Paste

End Sub


Private Sub barFerramentas_ButtonClick(ByVal Button As MSComctlLib.Button)
    'Botão  Negrito
    If Button.Index = 2 Then
        If TypeOf Me.ActiveControl Is RichTextBox Then
            Me.ActiveControl.SelBold = IIf(Me.ActiveControl.SelBold, False, True)
        End If
    End If
    'Botão  Itálico
    If Button.Index = 3 Then
        If TypeOf Me.ActiveControl Is RichTextBox Then
            Me.ActiveControl.SelItalic = IIf(Me.ActiveControl.SelItalic, False, True)
        End If
    End If
    'Botão  Sublinhado
    If Button.Index = 4 Then
        If TypeOf Me.ActiveControl Is RichTextBox Then
            Me.ActiveControl.SelUnderline = IIf(Me.ActiveControl.SelUnderline, False, True)
        End If
    End If
    'Botão Buscar Padrões
    If Button.Index = 6 Then
       If TypeOf Me.ActiveControl Is RichTextBox Then
            frmT46V144A.TxtAtualEspec = AtualEspec
            RchFinal.TextRTF = ""
            frmT46V144A.Show 1
            If RchFinal.Text <> "" Then
                 Me.ActiveControl.SelText = vbCrLf
                 Me.ActiveControl.SelRTF = RchFinal.TextRTF
            End If
        End If
    End If
    'Botao Visualizar
    If Button.Index = 8 Then
        'Verifica se tem dado a salvar.
        If frmT46V101A.mnuArqSalvar.Enabled = True Then
            If MsgBox("Para visualizar é necessário salvar toda a proposta!" & Chr(13) & "Deseja salvar a proposta nesse momento?", vbYesNo + vbQuestion) = vbNo Then
                Exit Sub
            End If
            If frmT46V101A.mfSalvar = False Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
        Call mpVisualizar(1, frmT46V101A.objstcPedido)
    End If
    'Botao imprimir
    If Button.Index = 9 Then
        'Verifica se tem dado a salvar.
        If frmT46V101A.mnuArqSalvar.Enabled = True Then
            If MsgBox("Para visualizar é necessário salvar toda a proposta!" & Chr(13) & "Deseja salvar a proposta nesse momento?", vbYesNo + vbQuestion) = vbNo Then
                Exit Sub
            End If
            If frmT46V101A.mfSalvar = False Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
        Call mpVisualizar(2, frmT46V101A.objstcPedido)
    End If
    'Botão Remover Especificação
    If Button.Index = 11 Then
        If MsgBox("Confirma exclusão da especificação?", vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
        End If
        Call gpLimPremios
        Unload Me
    End If
    'Botão Minimizar
    If Button.Index = 13 Then
        If TypeOf Me.ActiveControl Is RichTextBox Then
            If Me.ActiveControl.Tag <> "" Then
                Me.ActiveControl.Container.Top = Me.ActiveControl.Tag
                Me.ActiveControl.Tag = ""
                Me.ActiveControl.Container.Height = 735
                Me.ActiveControl.Height = 510
            End If
        End If
    End If
    'Botão Maximizar
    If Button.Index = 14 Then
        If TypeOf Me.ActiveControl Is RichTextBox Then
            If Me.ActiveControl.Tag = "" Then
                VScroll1 = 0
                Me.ActiveControl.Tag = Me.ActiveControl.Container.Top
                Me.ActiveControl.Container.Top = PicPanel.Top - 300
                Me.ActiveControl.Container.Height = PicPanel.Height - 130
                Me.ActiveControl.Height = PicPanel.Height - 400
                Me.ActiveControl.Container.ZOrder (0)
            End If
        End If
    End If
    'Botão Sair
    If Button.Index = 16 Then
        'Primeiro vai minimizar se tiver algum maximizado
        If TypeOf Me.ActiveControl Is RichTextBox Then
            If Me.ActiveControl.Tag <> "" Then
                Me.ActiveControl.Container.Top = Me.ActiveControl.Tag
                Me.ActiveControl.Tag = ""
                Me.ActiveControl.Container.Height = 735
                Me.ActiveControl.Height = 510
            End If
        End If
        Me.Hide
    End If
End Sub

Private Sub mpVisualizar(ByVal TypeShow As Integer, pObjstc As stcA46V702B)
    Dim lobjstcCorretor         As stcA46V705B      'Objeto da classe de estrutura stcA46V705B (TAB_PED_CORR)
    Dim lobjstcLoc              As stcA46V708B      'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC)
    Dim lobjstcSeguro           As stcA46V716B      'Objeto da classe de estrutura stcA46V716B (TAB_PED_LOC_SEGURO).
    Dim lobjstcCoss             As stcA46V706B      'Objeto da classe de estrutura stcA46V706B (TAB_PED_COSS).
    
    Dim NomeTemp As String
    Dim MyPath   As String
    Dim Aux      As String
    Dim i        As Integer
    Dim lCEP     As String
    
    On Error GoTo objError
    
    RchFinal.TextRTF = ""
    RchFinal.SelText = vbCrLf

    Call gpObtModalidade(pObjstc.Cod_Estip, pObjstc.Cod_Ramo)
    
    Call mpJuntarTextos("N° APÓLICE:", pObjstc.Num_Apol, 0, 3)
    Call mpJuntarTextos("MODALIDADE:", gModalidade, 0, 3)
    Call mpJuntarTextos("SEGURADO:", pObjstc.Nom_Segurado, 0, 3)
    
    RchFinal.SelText = vbCrLf
    
    If pObjstc.ColecaoPedCorr.Count > 0 Then
        Set lobjstcCorretor = pObjstc.ColecaoPedCorr.Item(1)
    Else
        Set lobjstcCorretor = New stcA46V705B
    End If
    
    If pObjstc.ColecaoPedLoc.Count > 0 Then
        Set lobjstcLoc = pObjstc.ColecaoPedLoc.Item(1)
    Else
        Set lobjstcLoc = New stcA46V708B
    End If
   
    Call mpJuntarTextos("CORRETOR:", lobjstcCorretor.gNom_Corretor, 0, 3)
    Call mpJuntarTextos("LIMITE MÁXIMO GARANTIA:", Format(pObjstc.Val_IS, "standard"), 0, 1)
    Call mpJuntarTextos("VALOR EM RISCO:", Format(lobjstcLoc.Val_Vr, "standard"), 0, 2)
    
    RchFinal.SelText = vbCrLf
    
    Call mpJuntarTextoRTF("OBJETO SEGURADO", pObjstc.objstcEspec.Dsc_Objetos)
    Call mpJuntarTextoRTF("RISCOS COBERTOS", pObjstc.objstcEspec.Dsc_RCobertos)
    
    'Local de Risco
    Aux = Trim(lobjstcLoc.Cod_Tip_Logr_Risco) & " "
    Aux = Aux & Trim(lobjstcLoc.Nom_Logr_Risco) & " "
    Aux = Aux & Trim(lobjstcLoc.objstcPedInsp.Num_Logr_Risco) & " - "
    Aux = Aux & Trim(lobjstcLoc.Dsc_Compl_Risco) & " - "
    Aux = Aux & Trim(lobjstcLoc.Nom_Bairro_Risco) & " - "
    Aux = Aux & Trim(lobjstcLoc.Nom_Cid_Risco) & " - "
    Aux = Aux & Trim(lobjstcLoc.Sig_UF_Risco) & ", Cep:"
    lCEP = Format(lobjstcLoc.Num_CEP_Risco, "00000000")
    Aux = Aux & Left(lCEP, 5) & "-" & Right(lCEP, 3) & " "
    
    RchFinal.SelText = vbCrLf
    Call mpJuntarTextos("LOCAL DE RISCO", Aux, 1, 0)
    RchFinal.SelText = vbCrLf
    
    'Endereço do segurado .
    Aux = Trim(pObjstc.objstcPedEndereco.Cod_Tip_Logr) & " "
    Aux = Aux & Trim(pObjstc.objstcPedEndereco.Nom_Logr) & " "
    Aux = Aux & Trim(pObjstc.objstcPedEndereco.Dsc_Compl_Logr) & " - "
    Aux = Aux & Trim(pObjstc.objstcPedEndereco.Nom_Bairro) & " - "
    Aux = Aux & Trim(pObjstc.objstcPedEndereco.Nom_Cid) & " - "
    Aux = Aux & Trim(pObjstc.objstcPedEndereco.Sig_UF) & ", Cep:"
    lCEP = Format(pObjstc.objstcPedEndereco.Num_Cep, "00000000")
    Aux = Aux & Left(lCEP, 5) & "-" & Right(lCEP, 3) & " "
    
    RchFinal.SelText = vbCrLf
    Call mpJuntarTextos("ENDEREÇO DO SEGURADO", Aux, 1, 0)
    RchFinal.SelText = vbCrLf
    
    Call mpJuntarTextoRTF("CARACTERÍSTICAS", pObjstc.objstcEspec.Dsc_Caracteristica)
    
    If gfPreenchido(pObjstc.gDesc_RamoAtiv) Then
        RchFinal.SelText = vbCrLf
        Call mpJuntarTextos("OCUPAÇÃO", pObjstc.gDesc_RamoAtiv, 1, 0)
    End If
    
    Call mpJuntarTextoRTF("PERÍMETRO DE COBERTURA", pObjstc.objstcEspec.Dsc_Perimetro)
    
    If gfPreenchido(pObjstc.objstcEspec.Dsc_LocalExposicao) Then
        RchFinal.SelText = vbCrLf
        Call mpJuntarTextos("LOCAL DE EXPOSIÇÃO", pObjstc.objstcEspec.Dsc_LocalExposicao, 1, 0)
    End If
    
    Call mpJuntarTextoRTF("ESPECIFICAÇÃO DO SERVIÇO", pObjstc.objstcEspec.Dsc_Especificacao)
    Call mpJuntarTextoRTF("COBERTURAS", pObjstc.objstcEspec.Dsc_Coberturas)
    Call mpJuntarTextoRTF("FRANQUIA", pObjstc.objstcEspec.Dsc_Franquia)
    
    RchFinal.SelText = vbCrLf
    Call mpJuntarTextos("PRAZO DO SEGURO:", gfForData(2, pObjstc.Dat_Ini_Vig) & "   à   " & gfForData(2, pObjstc.Dat_Ter_Vig), 0, 2)
    
    If pObjstc.objstcEspec.Dat_Obra <> 0 Then
        Call mpJuntarTextos("PRAZO DA OBRA:", gfForData(2, pObjstc.objstcEspec.Dat_Obra), 0, 2)
    End If
    If pObjstc.objstcEspec.Dat_Manut <> 0 Then
        Call mpJuntarTextos("PRAZO DA MANUTENÇÃO:", gfForData(2, pObjstc.objstcEspec.Dat_Manut), 0, 1)
    End If
 
    
    RchFinal.SelText = vbCrLf
    If lobjstcLoc.objstcPedClau.Cod_Clau <> 0 Then
        Aux = Val(Left(lobjstcLoc.objstcPedClau.Descricao, 14))
        If Len(Aux) <= 11 Then
            Call gpForCPF(1, Aux, Aux)
        Else
            
            Call gpForCNPJ(1, Aux, Aux)
        End If
        Call mpJuntarTextos("BENEFICIÁRIO:", Trim(Mid(lobjstcLoc.objstcPedClau.Descricao, 15, 60)) & " - CPF nº" & Aux, 0, 3)
    End If
    
    If pObjstc.ColecaoPedCoss.Count > 0 Then
        RchFinal.SelText = vbCrLf
        Call mpJuntarTextos("COSSEGURO:", "", 0, 0)
        For Each lobjstcCoss In pObjstc.ColecaoPedCoss
            If lobjstcCoss.Cod_Coss <> 0 Then
                Call mpJuntarTextos("", lobjstcCoss.Nom_Seguradora & vbTab & Format(lobjstcCoss.Per_Coss, "standard") & " %", 0, 1)
            End If
        Next lobjstcCoss
    End If
    
    If gfPreenchido(pObjstc.objstcEspec.Nom_Contratante) Then
        RchFinal.SelText = vbCrLf
        Call mpJuntarTextos("CONTRATANTE:", pObjstc.objstcEspec.Nom_Contratante, 0, 2)
    End If
        
    
    Call mpJuntarTextoRTF("CLÁUSULAS", pObjstc.objstcEspec.Dsc_Clausulas)
    Call mpJuntarTextoRTF("CONDIÇÕES ESPECIAIS", pObjstc.objstcEspec.Dsc_Condicao)
    

    If lobjstcLoc.ColecaoPedSeguro.Count > 0 Then
        RchFinal.SelText = vbCrLf
        Call mpJuntarTextos("OUTROS SEGUROS:", "", 0, 0)
        For Each lobjstcSeguro In lobjstcLoc.ColecaoPedSeguro
            Call mpJuntarTextos("", lobjstcSeguro.Dsc_Outro_Seguro & " - Apólice: " & lobjstcSeguro.Num_Apol_Ant & " - Item: " & lobjstcSeguro.Num_Item_Ant, 0, 1)
        Next lobjstcSeguro
    End If
    
    RchFinal.SelText = vbCrLf
    
    
    Call mpJuntarTextoRTF("OUTRAS INFORMAÇÕES", pObjstc.objstcEspec.Dsc_OutrasInf)
    Call mpJuntarTextoRTF("OBSERVAÇÃO", pObjstc.objstcEspec.Dsc_Observacao)
    
    RchFinal.SelText = vbCrLf
    Call mpJuntarTextos("PRÊMIO LÍQUIDO:", Format(pObjstc.Val_Pr, "standard"), 0, 3)
    On Error Resume Next
    MyPath = "c:\temp\"
    MkDir MyPath
    On Error GoTo objError
    Call mpDeletarTemp(MyPath)

    NomeTemp = "#Tmp" & Format(Now(), "hh:mm:ss") & ".ytf"
    NomeTemp = MyPath & Replace(NomeTemp, ":", "")
    
    'Salvar arquivo Montado
    RchFinal.SaveFile NomeTemp

    Dim objDoc      As Word.Document
    Dim objSec      As Word.Section
    
    Set objWord = New Word.Application
    objWord.Application.DisplayAlerts = wdAlertsNone
    Set objDoc = objWord.Documents.Open(NomeTemp)
    
    With objDoc.Paragraphs(1).Range
        .InsertParagraphBefore
        .InsertBefore "Apólice"
        .Font.Size = 24
        .bold = True
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Font.Name = "Arial"
        .InsertBefore vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine
    End With

    objDoc.Paragraphs(10).Range.InsertBreak Type:=wdPageBreak
    
    For i = 1 To objDoc.Paragraphs.Count - 1
        If i > 10 Then
            objDoc.Paragraphs(i).Alignment = wdAlignParagraphJustify
        End If
    Next
    
    Set objSec = objDoc.Sections(1)
    On Error Resume Next
    Dim oShape As Word.Shape
    Set oShape = objSec.Headers(wdHeaderFooterPrimary).Shapes.AddPicture(FileName:=App.Path & "\logoYasuda.gif", LinkToFile:=False) ', Anchor:=objSec.Headers(wdHeaderFooterPrimary).Range)
    With oShape
      .Height = 30
      .Width = 144
      .Left = 0
      .Top = 0
    End With
    Err = 0
    On Error GoTo objError

    objSec.Headers(wdHeaderFooterPrimary).Range.Font.Name = "Arial"
    objSec.Headers(wdHeaderFooterPrimary).Range.Font.Size = 12
    objSec.Headers(wdHeaderFooterPrimary).Range.Font.Italic = True
    objSec.Headers(wdHeaderFooterPrimary).Range.InsertAfter vbCrLf & Space(20) & pObjstc.gDesc_Ramo
    objSec.Headers(wdHeaderFooterPrimary).Range.Paragraphs(2).Alignment = wdAlignParagraphRight
    objSec.Headers(wdHeaderFooterPrimary).Range.InsertAfter vbCrLf & Space(20) & Trim(gModalidade)
    objSec.Headers(wdHeaderFooterPrimary).Range.Paragraphs(3).Alignment = wdAlignParagraphRight
   
    objSec.Headers(wdHeaderFooterPrimary).Range.InsertAfter vbCrLf & Space(34) & "C.N.P.J.  60.405.925"
    objSec.Headers(wdHeaderFooterPrimary).Range.Paragraphs(4).Range.Font.Size = 7
    objSec.Headers(wdHeaderFooterPrimary).Range.Paragraphs(4).Range.Font.Italic = False
    objSec.Headers(wdHeaderFooterPrimary).Range.Paragraphs(4).Alignment = wdAlignParagraphLeft
    
    objSec.Headers(wdHeaderFooterPrimary).Range.InsertAfter vbCrLf
    objSec.Headers(wdHeaderFooterPrimary).Range.InsertAfter vbCrLf
    objSec.Headers(wdHeaderFooterPrimary).Shapes.AddLine(50, 85, 600, 85).Line.Weight = 1
    
    objSec.Footers(wdHeaderFooterPrimary).Shapes.AddLine 50, 750, 600, 750
    objSec.Footers(wdHeaderFooterPrimary).Range.Font.Name = "Arial"
    objSec.Footers(wdHeaderFooterPrimary).Range.Font.Size = 7
    objSec.Footers(wdHeaderFooterPrimary).Range.Text = "ESPECIFICAÇÃO" & Space(180) & "Apólice: " & pObjstc.Num_Apol
    
    objSec.Footers(wdHeaderFooterPrimary).PageNumbers.Add PageNumberAlignment:=wdAlignPageNumberCenter
    objSec.Footers(wdHeaderFooterPrimary).PageNumbers.ChapterPageSeparator = wdSeparatorEmDash
    
    objDoc.PageSetup.LeftMargin = 70
    objDoc.PageSetup.RightMargin = 70
    objDoc.PageSetup.TopMargin = 70
    objDoc.PageSetup.BottomMargin = 70
    objDoc.PageSetup.HeaderDistance = 25

   
    objDoc.Save
    If TypeShow = 2 Then    'Imprimir
        objWord.PrintOut
    Else                    ' Mostrar na tela
        objWord.Application.Caption = "Especificação"
        objWord.ActiveWindow.View.Zoom.PageFit = wdPageFitBestFit
'        objWord.Application.DisplayStatusBar = False
        objWord.Application.WindowState = wdWindowStateMaximize
        objDoc.Protect (wdAllowOnlyComments)
        objDoc.Save
        objWord.Application.Visible = True
        objDoc.Save
    End If
    Exit Sub

objError:

    If Err <> 429 Then
        MsgBox Str$(Err) & " " & Error$
        Exit Sub
    Else
        Resume Next
    End If
End Sub

    
    
    




Sub AddLinesToHeader()

Dim oHeader As HeaderFooter, MyRange As Range, oLine As Shape
System.Cursor = wdCursorWait

Set oHeader = ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary)
'Make sure the shape anchored to the final paragraph in the Header
Set MyRange = oHeader.Range.Paragraphs.Last.Range

Set oLine = oHeader.Shapes.AddLine(0, 0, 0, 0, Anchor:=MyRange)
With oLine
'    .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
'    .RelativeVerticalPosition = wdRelativeVerticalPositionPage
    .Width = CentimetersToPoints(1)
    .Left = CentimetersToPoints(0.5)
    .Top = CentimetersToPoints(21.76)
'    .LockAnchor = True
    '.WrapFormat.Type = wdWrapNone
    'It's a good idea to name shapes you add, so that if you want to hide them,
    'or whatever later, you can refer to them by name
'    .Name = "First Line"
    'Add any other parameters you need, eg .Line.Weight = etc
End With

Set oLine = oHeader.Shapes.AddLine(0, 0, 0, 0, Anchor:=MyRange)
With oLine
'    .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
'    .RelativeVerticalPosition = wdRelativeVerticalPositionPage
    .Width = CentimetersToPoints(1)
    .Left = CentimetersToPoints(0.5)
    .Top = CentimetersToPoints(22.08)
'    .LockAnchor = True
'    .WrapFormat.Type = wdWrapNone
'    .Name = "Second Line"
    'add other parameters you need, eg .Line.Weight = etc
End With

System.Cursor = wdCursorNormal


Dim objShp As Word.Shape
Dim objDoc As Word.Document

        Dim objRng As Word.Range
        Set objRng = objDoc.StoryRanges.Item(wdPrimaryHeaderStory)  '
'I 'm guessing here...
        Set objShp = objDoc.Shapes.AddPicture(FileName:="C:\My.gif", LinkToFile:=False, Anchor:=objRng)
        With objShp
            .Height = InchesToPoints(1)
            .Width = InchesToPoints(1)
            .Left = InchesToPoints(0.25)
            .Top = InchesToPoints(0.25)

            'Do I need to use relative positioning?

        End With
End Sub


Sub TryAgainII()
Dim oFooter As Word.HeaderFooter
Dim oShape As Word.Shape
Dim oRng As Word.Range
Dim oDoc As Document
Set oRng = ActiveDocument.Range
Set oDoc = ActiveDocument

Application.ScreenUpdating = False
With oRng
  'Go to the end of the document
  .Collapse wdCollapseEnd
  .InsertBreak wdSectionBreakNextPage
  'Unlink the new (last section since it is at the end of the document)
'Section
  'from the previous one
  'oDoc.Sections.Last.Footers(wdHeaderFooterPrimary).LinkToPrevious = False
  'Set the reference to the footer
  'Set oFooter = oDoc.Sections.Last.Footers(wdHeaderFooterPrimary)
End With
'Set oShape = oDoc.Shapes.AddPicture(FileName:="C:\My.gif", LinkToFile:=False, Anchor:=oFooter.Range)
With oShape
  .Height = InchesToPoints(1)
  .Width = InchesToPoints(1)
  .Left = InchesToPoints(0.25)
  .Top = InchesToPoints(0.25)
End With
End Sub

Private Sub Form_Load()
    '1. Posiciona form.
    Call gpCenForm(Me)
    VScroll1 = 0
    PicPanel.Top = 400
    VScroll1.Max = (PicPanel.Height) / 3
    AtualEspec = e_Tipo_Espec_Nenhuma
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Screen.MousePointer = vbDefault
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
Dim i As Integer
    VScroll1.Left = Me.Width - 345
    PicPanel.Width = Me.Width - 330
    For i = 0 To Me.Controls.Count - 1
        If TypeOf Me.Controls(i) Is RichTextBox Then
            Me.Controls(i).Width = PicPanel.Width - 170
        End If
        If TypeOf Me.Controls(i) Is Frame Then
            Me.Controls(i).Width = PicPanel.Width - 90
        End If
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call mpDeletarTemp("c:\temp\")
End Sub

Private Sub PicPanel_GotFocus()
SendKeys "{TAB}"
End Sub


Private Sub RchCaracteristica_Change()
Call gpLimPremios
End Sub

Private Sub RchCaracteristica_LostFocus()
    Call mpTrocarFonte(RchCaracteristica)
End Sub

Private Sub RchClausulas_Change()
Call gpLimPremios
End Sub

Private Sub RchClausulas_GotFocus()
    VScroll1 = VScroll1.Max
    AtualEspec = e_Tipo_Espec_Clausulas
    mpRetiraSelecao
End Sub

Private Sub RchClausulas_LostFocus()
    Call mpTrocarFonte(RchClausulas)
End Sub

Private Sub RchCoberturas_Change()
Call gpLimPremios
End Sub

Private Sub RchCoberturas_LostFocus()
    Call mpTrocarFonte(RchCoberturas)
End Sub

Private Sub RchCondicao_Change()
Call gpLimPremios
End Sub

Private Sub RchCondicao_GotFocus()
    VScroll1 = VScroll1.Max
    AtualEspec = e_Tipo_Espec_Condicao
    Call mpRetiraSelecao
End Sub

Private Sub RchCondicao_LostFocus()
    Call mpTrocarFonte(RchCondicao)
End Sub

Private Sub RchEspecificacao_Change()
Call gpLimPremios
End Sub

Private Sub RchEspecificacao_LostFocus()
    Call mpTrocarFonte(RchEspecificacao)
End Sub

Private Sub RchFranquia_Change()
Call gpLimPremios
End Sub

Private Sub RchFranquia_LostFocus()
    Call mpTrocarFonte(RchFranquia)
End Sub

Private Sub RchObjetos_Change()
    Call gpLimPremios
End Sub

Private Sub RchObjetos_LostFocus()
    Call mpTrocarFonte(RchObjetos)
End Sub

Sub mpTrocarFonte(ByRef RchTexto As RichTextBox)
    RchTexto.SelStart = 1
    RchTexto.SelLength = Len(RchTexto.Text)
    If RchTexto.SelFontName <> "MS Sans Serif" Then
        RchTexto.SelFontName = "MS Sans Serif"
    End If
    RchTexto.SelStart = 1
    RchTexto.SelLength = 0
End Sub
Private Sub RchObservacao_Change()
    Call gpLimPremios
End Sub

Private Sub RchObservacao_LostFocus()
    Call mpTrocarFonte(RchObservacao)
End Sub

Private Sub RchOutrasInf_Change()
Call gpLimPremios
End Sub

Private Sub RchOutrasInf_LostFocus()
    Call mpTrocarFonte(RchOutrasInf)
End Sub

Private Sub RchPerimetro_Change()
    Call gpLimPremios
End Sub

Private Sub RchPerimetro_LostFocus()
    Call mpTrocarFonte(RchPerimetro)
End Sub

Private Sub RchRCobertos_Change()
Call gpLimPremios
End Sub

Private Sub RchRCobertos_GotFocus()
    VScroll1 = 300
    AtualEspec = e_Tipo_Espec_RCobertos
    Call mpRetiraSelecao
End Sub

Private Sub RchObservacao_GotFocus()
    VScroll1 = VScroll1.Max
    AtualEspec = e_Tipo_Espec_Observacao
    Call mpRetiraSelecao
End Sub

Private Sub RchOutrasInf_GotFocus()
    VScroll1 = VScroll1.Max
    AtualEspec = e_Tipo_Espec_OutrasInf
    Call mpRetiraSelecao
End Sub

Private Sub RchCoberturas_GotFocus()
    AtualEspec = e_Tipo_Espec_Coberturas
    Call mpRetiraSelecao
End Sub
Private Sub RchEspecificacao_GotFocus()
    AtualEspec = e_Tipo_Espec_Especificacao
    Call mpRetiraSelecao
End Sub
Private Sub RchPerimetro_GotFocus()
    VScroll1 = VScroll1.Max / 2
    AtualEspec = e_Tipo_Espec_Perimetro
    Call mpRetiraSelecao
End Sub

Private Sub RchObjetos_GotFocus()
    VScroll1 = 300
    AtualEspec = e_Tipo_Espec_Objetos
    Call mpRetiraSelecao
End Sub
Private Sub RchCaracteristica_GotFocus()
    AtualEspec = e_Tipo_Espec_Caracteristica
    Call mpRetiraSelecao
End Sub

Sub mpRetiraSelecao()
Dim i As Integer
    'For i = 0 To Me.Controls.Count - 1
    '    If TypeOf Me.Controls(i) Is RichTextBox Then
    '        If Me.Controls(i).Name = Me.ActiveControl.Name Then
    '            Me.Controls(i).HideSelection = False
    '       Else
    '            Me.Controls(i).HideSelection = True
    '        End If
    '    End If
    'Next
    'Screen.MousePointer = vbDefault
    'Me.MousePointer = vbDefault
End Sub
Private Sub RchFranquia_GotFocus()
    AtualEspec = e_Tipo_Espec_Franquia
    Call mpRetiraSelecao
End Sub

Private Sub RchRCobertos_LostFocus()
    Call mpTrocarFonte(RchRCobertos)
End Sub

Private Sub TxtContratante_Change()
Call gpLimPremios
End Sub

Private Sub TxtContratante_GotFocus()
    AtualEspec = e_Tipo_Espec_Nenhuma
    VScroll1 = 0
    With TxtContratante
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub TxtContratante_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub


Private Sub TxtLocalExposicao_Change()
Call gpLimPremios
End Sub

Private Sub TxtLocalExposicao_GotFocus()
    AtualEspec = e_Tipo_Espec_Nenhuma
    VScroll1 = 0
    With TxtLocalExposicao
        .SelStart = 0
        .SelLength = .MaxLength
    End With

End Sub


Private Sub TxtLocalExposicao_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub TxtPrazoManut_Change()
    Call gpLimPremios
End Sub

Private Sub TxtPrazoManut_GotFocus()
    AtualEspec = e_Tipo_Espec_Nenhuma
    VScroll1 = 0
    With TxtPrazoManut
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub TxtPrazoManut_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub


Private Sub TxtPrazoManut_LostFocus()
    TxtPrazoManut.Text = gfForData(9, TxtPrazoManut.Text)
    If IsDate(TxtPrazoManut.Text) = False Then
        TxtPrazoManut.Text = ""
    End If
End Sub

Private Sub TxtPrazoObra_Change()
Call gpLimPremios
End Sub

Private Sub TxtPrazoObra_GotFocus()
    AtualEspec = e_Tipo_Espec_Nenhuma
    VScroll1 = 0
    With TxtPrazoObra
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub


Private Sub TxtPrazoObra_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub


Private Sub TxtPrazoObra_LostFocus()
    TxtPrazoObra.Text = gfForData(9, TxtPrazoObra.Text)
    If IsDate(TxtPrazoObra.Text) = False Then
        TxtPrazoObra.Text = ""
    End If
End Sub

Private Sub VScroll1_Change()
    VScroll1_Scroll
End Sub


Private Sub VScroll1_Scroll()
    If VScroll1 < 400 Then
        PicPanel.Top = 400
    Else
        PicPanel.Top = -VScroll1
    End If
End Sub


