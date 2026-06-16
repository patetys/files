VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmPDF 
   Caption         =   "Exibindo arquivos PDF"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8145
   LinkTopic       =   "Form1"
   ScaleHeight     =   5460
   ScaleWidth      =   8145
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser wbrs1 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      ExtentX         =   11033
      ExtentY         =   8070
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmPDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'CyberTools - Projeto: Comum/Veículo e Perfil - Ari - Novembro\2013
'Public lsTituloForm As String

'dbianco - Implementação para o Caminho do relatório a ser exibido.
Option Explicit

Private strCaminhoRelatorio As String
Property Get CaminhoRelatorio() As String
    CaminhoRelatorio = strCaminhoRelatorio
End Property
Property Let CaminhoRelatorio(pCaminho As String)
    strCaminhoRelatorio = pCaminho
End Property

Public Sub gpIniciaForm(ByVal sTituloForm As String)
    Me.Caption = "Visualização do Relatório: " & sTituloForm
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    gpIniciaForm CaminhoRelatorio

    'define o tamanho do formulário e do controle WebBrowser Maximizados
    frmPDF.WindowState = 2    'Maximizado
    wbrs1.MenuBar = True
    
    'exibe o arquivo pdf selecionado no formulário anterior
    wbrs1.Navigate CaminhoRelatorio
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Resize()
    If Me.Width > 300 And Me.Height > 600 Then
        wbrs1.Top = 0
        wbrs1.Left = 0
        wbrs1.Width = Me.Width - 300
        wbrs1.Height = Me.Height - 600
    End If
End Sub
