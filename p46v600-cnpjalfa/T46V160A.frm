VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmT46V160A 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7230
   ClientLeft      =   1860
   ClientTop       =   1545
   ClientWidth     =   8520
   Icon            =   "T46V160A.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   8520
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser wbInternet 
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      ExtentX         =   15055
      ExtentY         =   12726
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
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   255
      Left            =   6600
      TabIndex        =   1
      Top             =   5160
      Width           =   495
   End
End
Attribute VB_Name = "frmT46V160A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mURL As String
Private mFecharForm As Boolean
Private mTelaOk As Boolean

Public Sub gpInstanciarForm(pURL As String, pFecharForm As Boolean, ByRef pTelaOk As Boolean)
    mURL = pURL
    mFecharForm = pFecharForm




    Dim lUrl As String
    Dim lText As String
    Dim lNumPag As Integer
    Dim lEndPag As String


    lEndPag = ""
    lText = "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">"
    lText = lText & " <HTML> <HEAD> <META name=""GENERATOR"" content=""IBM WebSphere Studio"">"
    lText = lText & " </HEAD> <BODY> <form name=""form1"" method=""POST"" action=""http://sp310114psluged/webworkflow/CONTENT/Custom_Prop_v5/fechar.asp"" >"
    lText = lText & " <script language=""JavaScript"">"
    lText = lText & "   window.open('" & mURL & "','','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes');"
    lText = lText & "   document.form1.submit();"
    lText = lText & " </script> "
    lText = lText & " </form>  </BODY> </HTML>"

    lNumPag = FreeFile
    lEndPag = "L:\User\AbrirPagina.html"
    If Dir(lEndPag) <> "" Then
        Kill lEndPag
    End If
    Open lEndPag For Output As #lNumPag
    Print #lNumPag, lText
    Close #lNumPag


    Screen.MousePointer = vbHourglass
    wbInternet.Navigate2 lEndPag
    pTelaOk = False


    'Screen.MousePointer = vbHourglass
    'wbInternet.Navigate2 mURL
    'pTelaOk = False

    'If mFecharForm = False Then
    Me.Show  'vbModal
    'End If
    pTelaOk = mTelaOk
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Resize()
    wbInternet.Height = Me.Height - 630
    wbInternet.Width = Me.Width - 345
End Sub

Private Sub wbInternet_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    If UCase(URL) Like "*/FECHAR.ASP*" Then
        mTelaOk = True
        Unload Me
    Else
        mTelaOk = False
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub wbInternet_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
'MsgBox URL
End Sub

Private Sub wbInternet_OnQuit()
    mTelaOk = True
    Unload Me
End Sub

Private Sub wbInternet_WindowClosing(ByVal IsChildWindow As Boolean, Cancel As Boolean)
    mTelaOk = True
    Unload Me
End Sub

