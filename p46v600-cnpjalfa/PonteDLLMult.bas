Attribute VB_Name = "BasPonteDLLMult"
Option Explicit
' ======================================================================================================================================================
' Name: BasPonteDLLMult
' Sistema: SyasMult
' Analista: Marcio Oliveira
'
'Objetivo: Simula o uso de variaveis contidas em classes como sendo de conteudo Global.
'          Foi implementado devido a migracao dos modulos para classes.
'          Os modulos continham variaveis, funcoes e procedures globais.
'
'*****Nao criar nenhum codigo novo nesta secao!!!!!
' ======================================================================================================================================================

Public Property Get gM46O001() As clsM46O001
    Static lsClsM46O001 As clsM46O001
    If lsClsM46O001 Is Nothing Then
        Set lsClsM46O001 = New clsM46O001
    End If

Retorno:
    Set gM46O001 = lsClsM46O001
End Property

Public Property Get gM46V999() As clsM46V999
    Static lsClsM46V999 As clsM46V999
    If lsClsM46V999 Is Nothing Then
        Set lsClsM46V999 = New clsM46V999
        Set lsClsM46V999.M46O001 = gM46O001
        Set lsClsM46V999.M46V111 = gM46V111
        Set lsClsM46V999.M46V103 = gM46V103
        Set lsClsM46V999.DynaTrace = New clsMDynaTrace
    End If

Retorno:
    Set gM46V999 = lsClsM46V999
End Property

Public Property Get gM46V998() As clsM46V998
    Static lsClsM46V998 As clsM46V998
    If lsClsM46V998 Is Nothing Then
        Set lsClsM46V998 = New clsM46V998
        Set lsClsM46V998.M46V999 = gM46V999
    End If

Retorno:
    Set gM46V998 = lsClsM46V998
End Property

Public Property Get gM46V111() As clsM46V111
    Static lsClsM46V111 As clsM46V111
    If lsClsM46V111 Is Nothing Then
        Set lsClsM46V111 = New clsM46V111
        Set lsClsM46V111.M46V999 = gM46V999
        Set lsClsM46V111.M99V001 = gM99V001
        Set lsClsM46V111.M46V103 = gM46V103
    End If

Retorno:
    Set gM46V111 = lsClsM46V111
End Property

Public Property Get gM46V001() As clsM46V001
    Static lsClsM46V001 As clsM46V001
    If lsClsM46V001 Is Nothing Then
        Set lsClsM46V001 = New clsM46V001
        Set lsClsM46V001.M46V111 = gM46V111
    End If

Retorno:
    Set gM46V001 = lsClsM46V001
End Property

Public Property Get gM46V103() As clsM46V103
    Static lsClsM46V103 As clsM46V103
    If lsClsM46V103 Is Nothing Then
        Set lsClsM46V103 = New clsM46V103
        Set lsClsM46V103.M46V111 = gM46V111
    End If

Retorno:
    Set gM46V103 = lsClsM46V103
End Property

Public Property Get gM99V001() As clsM99V001
    Static lsClsM99V001 As clsM99V001
    If lsClsM99V001 Is Nothing Then
        Set lsClsM99V001 = New clsM99V001
        Set lsClsM99V001.M46V999 = gM46V999
    End If

Retorno:
    Set gM99V001 = lsClsM99V001
End Property
