VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA46V750A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: clsA46V750A - Tab_Ped_Obs_SISCOTA
' Data/Hora: 19/10/2017
' Sistema: Multi
' Juliano Iwai
' ======================================================================================================================================================

Public ConexaoBD As ADODB.Connection    ' Conexão com o Banco de Dados
Public UltimoErro As String             ' Descrição do último erro gerado

Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "A46V750A"
End Property

Public Function mfFormatarInclusaoSQL(ByRef pobjstc As stcA46V750B, _
                                       ByVal pLogin As String) As String
    Dim lInsert As String 'Query.
    Dim pobjstcAnexo As stcA46V752B
    
    '' Novo
    lInsert = "Declare @IdLog integer " & vbCrLf
    
    lInsert = lInsert & "INSERT INTO RamosDiversos..Tab_Ped_SISCOTA_Obs "
    lInsert = lInsert & "(IdSiscota, Status, CodOrigem, Login, Observacao) VALUES ("
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.IdSiscota) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.STATUS) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.CodOrigem) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Login) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Observacao) & ")" & vbCrLf & vbCrLf
    
    lInsert = lInsert & "Set @IdLog=scope_identity()" & vbCrLf
    lInsert = lInsert & "INSERT INTO RamosDiversos..Tab_Ped_SISCOTA_Obs_LOG "
    lInsert = lInsert & "(IdObs, Data, Tipo, Descricao) VALUES ("
    lInsert = lInsert & "@IdLog, getdate(), 1, 'SISCOTA_Obs Inserido', "
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pLogin) & ")" & vbCrLf
    
    If (pobjstc.ColecaoAnexo.Count > 0) Then
        For Each pobjstcAnexo In pobjstc.ColecaoAnexo
            lInsert = lInsert & "INSERT INTO RamosDiversos..Tab_Ped_SISCOTA_Obs_ANEXO "
            lInsert = lInsert & "(IdObs, DataAnexo, NomeArquivo, PastaArquivo, NomeArquivoFisico) VALUES ("
            lInsert = lInsert & "@IdLog, getdate(), "
            lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstcAnexo.NomeArquivo) & ", "
            lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstcAnexo.PastaArquivo) & ", "
            lInsert = lInsert & "'" & pobjstcAnexo.NomeArquivoFisico & "') " & vbCrLf
        Next
    End If
    
    mfFormatarInclusaoSQL = lInsert
End Function

Public Function mfFormatarAtualizacaoSQL(ByRef pobjstc As stcA46V750B, _
                                       ByVal pLogin As String) As String
    Dim lUpdate As String 'Query.
    Dim pobjstcAnexo As stcA46V752B
    
    If pobjstc.IdObs > 0 Then
        '' Atualizar
        lUpdate = lUpdate & "UPDATE RamosDiversos..Tab_Ped_SISCOTA_Obs SET " & vbCrLf
        lUpdate = lUpdate & "IdSiscota=" & M46V999.gfFormatarCamposSQL(pobjstc.IdSiscota) & ","
        lUpdate = lUpdate & "Status=" & M46V999.gfFormatarCamposSQL(pobjstc.STATUS) & ","
        lUpdate = lUpdate & "CodOrigem=" & M46V999.gfFormatarCamposSQL(pobjstc.CodOrigem) & ","
        lUpdate = lUpdate & "Login=" & pLogin & ","
        lUpdate = lUpdate & "Observacao='" & pobjstc.Observacao & "' "
        lUpdate = lUpdate & " WHERE IdObs=" & M46V999.gfFormatarCamposSQL(pobjstc.IdObs) & vbCrLf & vbCrLf
        
        '' O correto seria colocar na classe 751, mas ficara aqui para dar agilidade
        lUpdate = lUpdate & "INSERT INTO RamosDiversos..Tab_Ped_SISCOTA_Obs_LOG "
        lUpdate = lUpdate & "(IdObs, Data, Tipo, Descricao) VALUES ("
        lUpdate = lUpdate & M46V999.gfFormatarCamposSQL(pobjstc.IdObs) & ", getdate(), 2, 'SISCOTA_Obs Alterado', "
        lUpdate = lUpdate & M46V999.gfFormatarCamposSQL(pLogin) & ")" & vbCrLf
        
'        If (pobjstc.ColecaoAnexo.Count > 0) Then
'            '' Remove todos e reinsere.
'            lUpdate = lUpdate & "DELETE FROM RamosDiversos..Tab_Ped_SISCOTA_ANEXO where IdObs=" & M46V999.gfFormatarCamposSQL(pobjstc.IdObs) & vbCrLf
'
'            For Each pobjstcAnexo In pobjstc.ColecaoAnexo
'                lUpdate = lUpdate & "INSERT INTO RamosDiversos..Tab_Ped_SISCOTA_ANEXO "
'                lUpdate = lUpdate & "(IdObs, DataAnexo, NomeArquivo, PastaArquivo) VALUES ("
'                lUpdate = lUpdate & M46V999.gfFormatarCamposSQL(pobjstcAnexo.IdObs) & ", "
'                lUpdate = lUpdate & M46V999.gfFormatarCamposSQL(pobjstcAnexo.DataAnexo) & ", "
'                lUpdate = lUpdate & M46V999.gfFormatarCamposSQL(pobjstcAnexo.NomeArquivo) & ", "
'                lUpdate = lUpdate & M46V999.gfFormatarCamposSQL(pobjstcAnexo.PastaArquivo) & ") " & vbCrLf
'            Next
'        End If
    Else
        lUpdate = mfFormatarInclusaoSQL(pobjstc, pLogin)
    End If
    
    mfFormatarAtualizacaoSQL = lUpdate
End Function

Public Function mfFormatarExclusaoSQL(ByRef pobjstc As stcA46V750B, _
                                       ByVal pLogin As String) As String
    Dim lDelete As String 'Query
        
    If pobjstc.IdObs > 0 Then
        '' Remocao Logica: status=2
        lDelete = "UPDATE RamosDiversos..Tab_Ped_SISCOTA_Obs SET "
        lDelete = lDelete & "Status = 2 where IdObs = " & M46V999.gfFormatarCamposSQL(pobjstc.IdObs) & vbCrLf
        
        '' O correto seria colocar na classe 751A, mas ficará aqui para simplificar
        lDelete = lDelete & "INSERT INTO RamosDiversos..Tab_Ped_SISCOTA_Obs_LOG "
        lDelete = lDelete & "(IdObs, Data, Tipo, Descricao) VALUES ("
        lDelete = lDelete & M46V999.gfFormatarCamposSQL(pobjstc.IdObs) & ", getdate(), 3, 'SISCOTA_Obs Removido')" & vbCrLf
    End If
    
    mfFormatarExclusaoSQL = lDelete
End Function

Public Function mfFormatarConsultaSQL(ByRef pobjstc As stcA46V750B) As String
    Dim lSelect As String 'Query.
    
    lSelect = "SELECT TOP 1 IdObs, IdSiscota, Status, CodOrigem, Login, Observacao FROM RamosDiversos..Tab_Ped_SISCOTA_Obs (nolock) Where "
    lSelect = lSelect & " IdObs = " & M46V999.gfFormatarCamposSQL(pobjstc.IdObs)
    lSelect = lSelect & " Order by IdObs DESC"

    mfFormatarConsultaSQL = lSelect
End Function

Public Function mfFormatarConsultaSQLByIdSiscota(ByVal IdSiscota As Long) As String
    Dim lSelect As String 'Query.
    
    lSelect = "SELECT IdObs, IdSiscota, Status, CodOrigem, Login, Observacao FROM RamosDiversos..Tab_Ped_SISCOTA_Obs (nolock) Where "
    lSelect = lSelect & " IdSiscota = " & M46V999.gfFormatarCamposSQL(IdSiscota)
    lSelect = lSelect & " Order by IdObs DESC"

    mfFormatarConsultaSQLByIdSiscota = lSelect
End Function

Public Sub mpCarregarAtributos(ByRef pobjstc As stcA46V750B, prst As clsYasRecordSet)
    pobjstc.IdObs = M46V999.gfTratarNulo(prst("IdObs"), tpoNumerico)
    pobjstc.IdSiscota = M46V999.gfTratarNulo(prst("IdSiscota"), tpoNumerico)
    pobjstc.STATUS = M46V999.gfTratarNulo(prst("Status"), tpoNumerico)
    pobjstc.CodOrigem = M46V999.gfTratarNulo(prst("CodOrigem"), tpoString)
    pobjstc.Login = M46V999.gfTratarNulo(prst("Login"), tpoNumerico)
    pobjstc.Observacao = M46V999.gfTratarNulo(prst("Observacao"), tpoString)
End Sub

Public Function mfCarregar(ByRef pobjstc As stcA46V750B, ByRef Mensagem As String) As Boolean
    Dim lrstObs As clsYasRecordSet      'Registro.
    Dim lrstLog As clsYasRecordSet      'Registro.
    Dim lrstAnexo As clsYasRecordSet    'Registro.
    
    Dim objLog As stcA46V751B
    Dim objAnexo As stcA46V752B
    Dim lsSQl As String                 'SQL.

    On Error GoTo ControleErros
    
    If Not mfConectar Then GoTo ControleErros
    
    lsSQl = mfFormatarConsultaSQL(pobjstc)
    
    If M46V999.gfObtRegistro(ConexaoBD, lsSQl, lrstObs, UltimoErro) = False Then GoTo ControleErros
    
    If Not lrstObs.EOF Then
        Call mpCarregarAtributos(pobjstc, lrstObs)
    Else
        Set lrstObs = Nothing
        Mensagem = "Registro não localizado com Id=" & M46V999.gfTratarNulo(pobjstc.IdObs, tpoNumerico)
        GoTo Fim
    End If
    
    '' Se existe Obs_SISCOTA, carregar LOGs e Anexos
    If (Not pobjstc Is Nothing) Then
        '' Carrega LOGs
        Dim objLogA As clsA46V751A
        Set objLogA = New clsA46V751A
        Set objLogA.M46V999 = Me.M46V999
        
        lsSQl = objLogA.mfFormatarConsultaSQL(pobjstc.IdObs)
        
        If M46V999.gfObtRegistro(ConexaoBD, lsSQl, lrstLog, UltimoErro) = False Then GoTo Fim
        
        Do While Not lrstLog.EOF
            Set objLog = New stcA46V751B
            Call objLogA.mpCarregarAtributos(objLog, lrstLog)
               
            pobjstc.ColecaoLog.Add objLog
            
            lrstLog.MoveNext
        Loop
        
        Set lrstLog = Nothing
                
        '' Carrega Anexos
        Dim objAnexoA As clsA46V752A
        Set objAnexoA = New clsA46V752A
        Set objAnexoA.M46V999 = Me.M46V999
        
        lsSQl = objAnexoA.mfFormatarConsultaSQL(pobjstc.IdObs)
        
        If M46V999.gfObtRegistro(ConexaoBD, lsSQl, lrstAnexo, UltimoErro) = False Then GoTo Fim
        
        Do While Not lrstAnexo.EOF
            Set objAnexo = New stcA46V752B
            Call objAnexoA.mpCarregarAtributos(objAnexo, lrstAnexo)
                        
            pobjstc.ColecaoAnexo.Add objAnexo
            
            lrstAnexo.MoveNext
        Loop
        
        Set lrstAnexo = Nothing
    End If
    
Fim:
    '' Fim com sucesso
    mfCarregar = True
    Exit Function

ControleErros:
    Set lrstObs = Nothing
    mfCarregar = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V750B.mfCarregar ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V750B.mfCarregar ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
End Function

Public Function mfSalvar(ByRef pobjstc As stcA46V750B, _
                         ByVal pLogin As String, _
                         ByRef Mensagem As String) As Boolean

    Dim lsSQl As String           'Query.
    
    Mensagem = ""
    
    On Local Error GoTo ControleErros
    
    lsSQl = "mfConectar"
    If Not mfConectar Then GoTo ControleErros
    
    If pobjstc.IdObs > 0 Then
        lsSQl = mfFormatarAtualizacaoSQL(pobjstc, pLogin)
    Else
        lsSQl = mfFormatarInclusaoSQL(pobjstc, pLogin)
    End If
    
    If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then
        GoTo ControleErros
    End If
    
    mfSalvar = True
    Exit Function
    
ControleErros:
    Mensagem = "A46V750A.mfIncluir ERRO - " & UltimoErro & vbCrLf & " SQL: " & lsSQl
    Call M46V999.gpGraLog(0, Mensagem)

End Function


Public Function mfExcluir(ByRef pobjstc As stcA46V750B, _
                          ByVal pLogin As String, _
                          ByRef Mensagem As String) As Boolean
    
    Dim lsSQl As String           'Query.
    
    Mensagem = ""
    
    On Local Error GoTo ControleErros
    
    lsSQl = "mfConectar"
    If Not mfConectar Then GoTo ControleErros

    lsSQl = mfFormatarExclusaoSQL(pobjstc, pLogin)
    
    If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then
        GoTo ControleErros
    End If
    
    mfExcluir = True
    Exit Function

ControleErros:
    Mensagem = "A46V750A.mfIncluir ERRO - " & UltimoErro & vbCrLf & " SQL: " & lsSQl
    Call M46V999.gpGraLog(0, Mensagem)

End Function

Private Sub Class_Initialize()
    Set mClsM46V999 = Nothing
End Sub

Private Sub Class_Terminate()
    Set mClsM46V999 = Nothing
End Sub

Private Function mfConectar() As Boolean
    mfConectar = False
    If ConexaoBD Is Nothing Then
        If M46V999.gfAbrBasDados("P00MPED", M46V999.gTipP00MPED, M46V999.gEndP00MPED, ConexaoBD, _
                                 UltimoErro) = False Then Exit Function
    ElseIf ConexaoBD.State <> 1 Then
        If M46V999.gfAbrBasDados("P00MPED", M46V999.gTipP00MPED, M46V999.gEndP00MPED, ConexaoBD, _
                                 UltimoErro) = False Then Exit Function
    End If
    mfConectar = True
End Function

Public Property Set M46V999(pClsM46V999 As clsM46V999)
    Set mClsM46V999 = pClsM46V999
End Property

Public Property Get M46V999() As clsM46V999
    If mClsM46V999 Is Nothing Then
        Err.Raise 91, Me.Class_Name, "Objeto M46V999 não atribuído para a classe (" & Me.Class_Name & ")"
        '***Esta Classe DEVE utilizar uma instancia da classe clsM46V999 já
        '   iniciada no escopo onde esta classe foi instanciada.
        'Atribua a o objeto (clsM46V999 ja instanciado) no momento que eu inicializou esta classe (atraves do Property SET)!!!
    End If

Retorno:
    Set M46V999 = mClsM46V999
End Property


