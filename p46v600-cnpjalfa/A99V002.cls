VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA99V002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Constantes que podem ser especificadas no membro ulFlags da estrutura BROWSEINFO.
'Apenas retornar diretório de arquivo.
'Se usuário selecionar pastas que não são parte de arquivo, o botão OK está desabilitado.
Private Const BIF_RETURNONLYFSDIRS = &H1

Private Type BROWSEINFO
    hOwner          As Long
    pidlRoot        As Long
    pszDisplayName  As String
    lpszTitle       As String
    ulFlags         As Long
    lpfn            As Long
    lParam          As Long
    iImage          As Long
End Type

'Diálogo de busca.
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" _
                                           (lpBrowseInfo As BROWSEINFO) As Long

'Posiciona nome do caminho da estrutura de retorno.
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" _
                                             (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Function gfObtDiretorio(ByVal pMensagem As String) As String
'Função: obtém diretório.

    Dim lDiretorio      As String       'Diretório - retorno.
    Dim lDirSelecionado As Long         'Diretório selecionado.
    Dim udtBrowseInfo   As BROWSEINFO

    'Especifica janela para o quadro de diálogo.
    udtBrowseInfo.hOwner = 0&

    'Especifica o início da busca (nulo indica Meu Computador).
    'udtBrowseInfo.pidlRoot = 0&
    udtBrowseInfo.pidlRoot = 0

    'Posiciona mensagem.
    udtBrowseInfo.lpszTitle = pMensagem

    'Expecifica indicadores.
    udtBrowseInfo.ulFlags = BIF_RETURNONLYFSDIRS

    'Executa API.
    'O retorno indica o diretório selecionado.
    'Se usuário cancelar quadro de diálogo, o valor de retorno é zero.
    lDirSelecionado = SHBrowseForFolder(udtBrowseInfo)
    lDiretorio = Space$(512)
    If SHGetPathFromIDList(lDirSelecionado, lDiretorio) Then
        gfObtDiretorio = mpExcNulos(lDiretorio)
    End If
End Function
Private Function mpExcNulos(ByVal pCampo As String) As String
    'Elimina valores nulos do campo.

    Dim lPosicao    As Long     'Utilizado com InStr.

    mpExcNulos = pCampo
    lPosicao = InStr(pCampo, vbNullChar)
    If lPosicao > 0 Then
        mpExcNulos = Left$(mpExcNulos, lPosicao - 1)
    End If
End Function
