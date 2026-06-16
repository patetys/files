VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA46V753A"
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
    Class_Name = "A46V753A"
End Property

Public Function mfFormatarInclusaoSQL(ByRef pobjstcSis As stcA46V753B, _
                                      ByVal pLogin As String) As String
    Dim lInsert As String 'Query.
    Dim pobjstcAnexo As stcA46V752B
    Dim pobjstcObs As New stcA46V750B

    '' Novo
    lInsert = "Declare @IdSiscota integer, @IdObs integer " & vbCrLf
    lInsert = lInsert & "INSERT INTO RamosDiversos..Tab_Ped_SISCOTA "

    If pobjstcSis.IdGrupo = "" Then
        lInsert = lInsert & "(Nosso_Numero, Num_Ped, Siscota, Status) VALUES ("
    Else
        lInsert = lInsert & "(Nosso_Numero, Num_Ped, IdGrupo, Siscota, Status) VALUES ("
    End If

    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstcSis.nosso_numero) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstcSis.Num_Ped) & ","

    If pobjstcSis.IdGrupo <> "" Then
        lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstcSis.IdGrupo) & ","
    End If

    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstcSis.Siscota) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstcSis.STATUS) & ")" & vbCrLf & vbCrLf

    lInsert = lInsert & "Set @IdSiscota=scope_identity()" & vbCrLf
    lInsert = lInsert & "INSERT INTO RamosDiversos..Tab_Ped_SISCOTA_Obs "
    lInsert = lInsert & "(IdSiscota, Status, CodOrigem, Login, Observacao) VALUES ("
    lInsert = lInsert & "@IdSiscota, "
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstcSis.ColecaoObs(1).STATUS) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstcSis.ColecaoObs(1).CodOrigem) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstcSis.ColecaoObs(1).Login) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstcSis.ColecaoObs(1).Observacao) & ")" & vbCrLf & vbCrLf

    lInsert = lInsert & "Set @IdObs=@@identity" & vbCrLf
    lInsert = lInsert & "INSERT INTO RamosDiversos..Tab_Ped_SISCOTA_LOG "
    lInsert = lInsert & "(IdSiscota, Data, Tipo, Descricao) VALUES ("
    lInsert = lInsert & "@IdSiscota, getdate(), 1, 'SISCOTA inserido')" & vbCrLf

    lInsert = lInsert & "INSERT INTO RamosDiversos..Tab_Ped_SISCOTA_Obs_LOG "
    lInsert = lInsert & "(IdObs, Data, Tipo, Descricao) VALUES ("
    lInsert = lInsert & "@IdObs, getdate(), 1, 'SISCOTA Obs inserido')" & vbCrLf

    If pobjstcSis.ColecaoObs.Count > 0 Then
        Set pobjstcObs = pobjstcSis.ColecaoObs(1)
    End If

    If (pobjstcObs.ColecaoAnexo.Count > 0) Then
        For Each pobjstcAnexo In pobjstcObs.ColecaoAnexo
            lInsert = lInsert & "INSERT INTO RamosDiversos..Tab_Ped_SISCOTA_ANEXO "
            lInsert = lInsert & "(IdObs, NomeArquivo, PastaArquivo, NomeArquivoFisico) VALUES ("
            lInsert = lInsert & "@IdObs, "
            lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstcAnexo.NomeArquivo) & ", "
            lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstcAnexo.PastaArquivo) & ", "
            lInsert = lInsert & "'" & pobjstcAnexo.NomeArquivoFisico & "') " & vbCrLf
        Next
    End If

    mfFormatarInclusaoSQL = lInsert
End Function

Public Function mfFormatarInclusaoSQLObs(ByRef pobjstcObs As stcA46V750B, _
                                      ByVal pLogin As String) As String
    Dim lInsert As String 'Query.
    Dim pobjstcAnexo As stcA46V752B

    '' Novo
    lInsert = "Declare @IdObs integer " & vbCrLf
    lInsert = lInsert & "INSERT INTO RamosDiversos..Tab_Ped_SISCOTA_Obs "
    lInsert = lInsert & "(IdSiscota, Status, CodOrigem, Login, Observacao) VALUES ("
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstcObs.IdSiscota) & ", "
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstcObs.STATUS) & ", "
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstcObs.CodOrigem) & ", "
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstcObs.Login) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstcObs.Observacao) & ")" & vbCrLf & vbCrLf

    lInsert = lInsert & "Set @IdObs=@@identity" & vbCrLf
    lInsert = lInsert & "INSERT INTO RamosDiversos..Tab_Ped_SISCOTA_Obs_LOG "
    lInsert = lInsert & "(IdObs, Data, Tipo, Descricao) VALUES ("
    lInsert = lInsert & "@IdObs, getdate(), 1, 'SISCOTA Obs inserido')" & vbCrLf

    If (pobjstcObs.ColecaoAnexo.Count > 0) Then
        For Each pobjstcAnexo In pobjstcObs.ColecaoAnexo
            lInsert = lInsert & "INSERT INTO RamosDiversos..Tab_Ped_SISCOTA_ANEXO "
            lInsert = lInsert & "(IdObs, NomeArquivo, PastaArquivo, NomeArquivoFisico) VALUES ("
            lInsert = lInsert & "@IdObs, "
            lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstcAnexo.NomeArquivo) & ", "
            lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstcAnexo.PastaArquivo) & ", "
            lInsert = lInsert & "'" & pobjstcAnexo.NomeArquivoFisico & "') " & vbCrLf
        Next
    End If

    mfFormatarInclusaoSQLObs = lInsert
End Function

Public Function mfFormatarAtualizacaoSQLSis(ByRef pobjstcSis As stcA46V753B, _
                                         ByVal pLogin As String) As String
    Dim lUpdate As String 'Query.

    If pobjstcSis.IdSiscota > 0 Then
        '' Atualizar
        lUpdate = ""
        lUpdate = lUpdate & "UPDATE RamosDiversos..Tab_Ped_SISCOTA SET " & vbCrLf
        lUpdate = lUpdate & "NOSSO_NUMERO=" & M46V999.gfFormatarCamposSQL(pobjstcSis.nosso_numero) & ", "
        lUpdate = lUpdate & "Num_Ped=" & M46V999.gfFormatarCamposSQL(pobjstcSis.Num_Ped) & ", "
        lUpdate = lUpdate & "IdGrupo=" & M46V999.gfFormatarCamposSQL(pobjstcSis.IdGrupo) & ", "
        lUpdate = lUpdate & "SISCOTA=" & M46V999.gfFormatarCamposSQL(pobjstcSis.Siscota) & ", "
        lUpdate = lUpdate & "Status=" & M46V999.gfFormatarCamposSQL(pobjstcSis.STATUS)
        lUpdate = lUpdate & " WHERE IdSiscota=" & M46V999.gfFormatarCamposSQL(pobjstcSis.IdSiscota) & vbCrLf & vbCrLf

        '' O correto seria colocar na classe 751, mas ficara aqui para dar agilidade
        lUpdate = lUpdate & "INSERT INTO RamosDiversos..Tab_Ped_SISCOTA_LOG "
        lUpdate = lUpdate & "(IdSiscota, Data, Tipo, Descricao) VALUES ("
        lUpdate = lUpdate & M46V999.gfFormatarCamposSQL(pobjstcSis.IdSiscota) & ", getdate(), "
        lUpdate = lUpdate & M46V999.gfFormatarCamposSQL(pobjstcSis.STATUS) & ", 'SISCOTA Alterado') " & vbCrLf & vbCrLf
    End If

    mfFormatarAtualizacaoSQLSis = lUpdate
End Function

Public Function mfFormatarAtualizacaoSQLObs(ByRef pobjstcObs As stcA46V750B, _
                                         ByVal pLogin As String) As String
    Dim lUpdate As String 'Query.
    Dim pobjstcAnexo As stcA46V752B

    If pobjstcObs.IdSiscota > 0 Then
        '' Atualizar
        lUpdate = lUpdate & "UPDATE RamosDiversos..Tab_Ped_SISCOTA_Obs SET " & vbCrLf
        lUpdate = lUpdate & "IdSiscota=" & M46V999.gfFormatarCamposSQL(pobjstcObs.IdSiscota) & ", "
        lUpdate = lUpdate & "CodOrigem=" & M46V999.gfFormatarCamposSQL(pobjstcObs.CodOrigem) & ", "
        lUpdate = lUpdate & "Login=" & M46V999.gfFormatarCamposSQL(pobjstcObs.Login) & ", "
        lUpdate = lUpdate & "Status=" & M46V999.gfFormatarCamposSQL(pobjstcObs.STATUS) & ", "
        lUpdate = lUpdate & "Observacao='" & pobjstcObs.Observacao & "' "
        lUpdate = lUpdate & " WHERE IdObs=" & M46V999.gfFormatarCamposSQL(pobjstcObs.IdObs) & vbCrLf & vbCrLf

        lUpdate = lUpdate & "INSERT INTO RamosDiversos..Tab_Ped_SISCOTA_Obs_LOG "
        lUpdate = lUpdate & "(IdObs, Data, Tipo, Descricao) VALUES ("
        lUpdate = lUpdate & M46V999.gfFormatarCamposSQL(pobjstcObs.IdObs) & ", getdate(), "
        lUpdate = lUpdate & M46V999.gfFormatarCamposSQL(2) & ", 'SISCOTA_Obs Alterado') " & vbCrLf & vbCrLf

        '' Remove todos e reinsere.
        lUpdate = lUpdate & "DELETE FROM RamosDiversos..Tab_Ped_SISCOTA_ANEXO where IdObs=" & M46V999.gfFormatarCamposSQL(pobjstcObs.IdObs) & vbCrLf

        If (pobjstcObs.ColecaoAnexo.Count > 0) Then
            For Each pobjstcAnexo In pobjstcObs.ColecaoAnexo
                lUpdate = lUpdate & "INSERT INTO RamosDiversos..Tab_Ped_SISCOTA_ANEXO "
                lUpdate = lUpdate & "(IdObs, NomeArquivo, PastaArquivo, NomeArquivoFisico) VALUES ("
                lUpdate = lUpdate & M46V999.gfFormatarCamposSQL(pobjstcAnexo.IdObs) & ", "
                lUpdate = lUpdate & M46V999.gfFormatarCamposSQL(pobjstcAnexo.NomeArquivo) & ", "
                lUpdate = lUpdate & M46V999.gfFormatarCamposSQL(pobjstcAnexo.PastaArquivo) & ", "
                lUpdate = lUpdate & "'" & pobjstcAnexo.NomeArquivoFisico & "') " & vbCrLf

                '                lUpdate = lUpdate & "INSERT INTO RamosDiversos..Tab_Ped_SISCOTA_Obs_LOG "
                '                lUpdate = lUpdate & "(IdObs, Data, Tipo, Descricao) VALUES ("
                '                lUpdate = lUpdate & M46V999.gfFormatarCamposSQL(pobjstcAnexo.IdObs) & ", getdate(), "
                '                lUpdate = lUpdate & M46V999.gfFormatarCamposSQL(pobjstcObs.STATUS) & ", 'SISCOTA_Obs Anexos Alterados') " & vbCrLf & vbCrLf
            Next
            '        Else
            '            lUpdate = lUpdate & "DELETE FROM RamosDiversos..Tab_Ped_SISCOTA_ANEXO where IdObs=" & M46V999.gfFormatarCamposSQL(pobjstcObs.IdObs) & vbCrLf & vbCrLf
            '
            '            lUpdate = lUpdate & "INSERT INTO RamosDiversos..Tab_Ped_SISCOTA_Obs_LOG "
            '            lUpdate = lUpdate & "(IdObs, Data, Tipo, Descricao) VALUES ("
            '            lUpdate = lUpdate & M46V999.gfFormatarCamposSQL(pobjstcObs.IdObs) & ", getdate(), "
            '            lUpdate = lUpdate & M46V999.gfFormatarCamposSQL(pobjstcObs.STATUS) & ", 'SISCOTA_Obs Anexos Removidos') " & vbCrLf & vbCrLf
        End If
    End If

    mfFormatarAtualizacaoSQLObs = lUpdate
End Function

Public Function mfFormatarAtualizacaoCodOrigemObsSQL(ByRef pCodOrigem As Integer, _
                                                     ByRef pIdObs As Long, _
                                                     ByVal pLogin As String) As String
    Dim lUpdate As String 'Query.

    '' Atualizar
    lUpdate = lUpdate & "UPDATE RamosDiversos..Tab_Ped_SISCOTA_Obs SET " & vbCrLf
    lUpdate = lUpdate & "CodOrigem=" & M46V999.gfFormatarCamposSQL(pCodOrigem)
    lUpdate = lUpdate & " WHERE IdObs=" & M46V999.gfFormatarCamposSQL(pIdObs) & vbCrLf & vbCrLf

    lUpdate = lUpdate & "INSERT INTO RamosDiversos..Tab_Ped_SISCOTA_Obs_LOG "
    lUpdate = lUpdate & "(IdObs, Data, Tipo, Descricao) VALUES ("
    lUpdate = lUpdate & M46V999.gfFormatarCamposSQL(pIdObs) & ", getdate(), "
    lUpdate = lUpdate & M46V999.gfFormatarCamposSQL(2) & ", 'SISCOTA_Obs Alterado') " & vbCrLf & vbCrLf

    mfFormatarAtualizacaoCodOrigemObsSQL = lUpdate
End Function

Public Function mfFormatarConsultaSQL(ByRef pobjstcSis As stcA46V753B) As String
    Dim lSelect As String 'Query.

    lSelect = "SELECT TOP 1 RamosDiversos..Tab_Ped_SISCOTA.* FROM RamosDiversos..Tab_Ped_SISCOTA (nolock) "

    If (pobjstcSis.IdSiscota > 0) Then
        lSelect = lSelect & " Where IdSiscota = " & M46V999.gfFormatarCamposSQL(pobjstcSis.IdSiscota)
    ElseIf (pobjstcSis.ColecaoObs(1).IdObs > 0) Then
        lSelect = lSelect & " INNER JOIN  RAMOSDIVERSOS..TAB_PED_SISCOTA_OBS (nolock) ON RamosDiversos..Tab_Ped_SISCOTA.IdSiscota = RAMOSDIVERSOS..TAB_PED_SISCOTA_OBS.IdSiscota "
        lSelect = lSelect & " WHERE IdObs = " & M46V999.gfFormatarCamposSQL(pobjstcSis.ColecaoObs(1).IdObs)
    ElseIf (pobjstcSis.nosso_numero <> "") Then
        lSelect = lSelect & " WHERE Nosso_Numero = " & M46V999.gfFormatarCamposSQL(pobjstcSis.nosso_numero)
    Else
        lSelect = ""
    End If

    If lSelect <> "" Then
        lSelect = lSelect & " Order by IdObs DESC"
    End If

    mfFormatarConsultaSQL = lSelect
End Function

Public Function mfFormatarConsultaSQLNossoNumero(ByVal strNossoNumero As String) As String
    Dim lSelect As String 'Query.

    lSelect = "SELECT TOP 1 * FROM RamosDiversos..Tab_Ped_SISCOTA (nolock) Where Nosso_Numero = " & M46V999.gfFormatarCamposSQL(strNossoNumero)
    lSelect = lSelect & " Order by IdSiscota DESC"

    mfFormatarConsultaSQLNossoNumero = lSelect
End Function

Public Sub mpCarregarAtributos(ByRef pobjstc As stcA46V753B, prst As clsYasRecordSet)
    pobjstc.IdSiscota = M46V999.gfTratarNulo(prst("IdSiscota"), tpoNumerico)
    pobjstc.IdGrupo = M46V999.gfTratarNulo(prst("IdGrupo"), tpoNumerico)
    pobjstc.nosso_numero = M46V999.gfTratarNulo(prst("Nosso_Numero"), tpoString)
    pobjstc.Num_Ped = M46V999.gfTratarNulo(prst("Num_Ped"), tpoNumerico)
    pobjstc.Siscota = M46V999.gfTratarNulo(prst("SISCOTA"), tpoString)
    pobjstc.STATUS = M46V999.gfTratarNulo(prst("Status"), tpoNumerico)
End Sub

Public Function mfCarregarMinimo(ByRef pobjstcSis As stcA46V753B, ByRef Mensagem As String) As Boolean
    Dim lrstSis As clsYasRecordSet      'Registro.
    Dim lrstObs As clsYasRecordSet      'Registro.
    Dim lrstLog As clsYasRecordSet      'Registro.
    Dim lrstAnexo As clsYasRecordSet    'Registro.

    Dim objObs As stcA46V750B
    Dim objLog As stcA46V751B
    Dim objAnexo As stcA46V752B
    Dim lsSQl As String                 'SQL.

    On Error GoTo ControleErros

    If Not mfConectar Then GoTo ControleErros

    lsSQl = mfFormatarConsultaSQLNossoNumero(pobjstcSis.nosso_numero)

    If M46V999.gfObtRegistro(ConexaoBD, lsSQl, lrstSis, UltimoErro) = False Then GoTo ControleErros

    If Not lrstSis.EOF Then
        Call mpCarregarAtributos(pobjstcSis, lrstSis)
    Else
        Set lrstSis = Nothing
        GoTo Fim
    End If

    '    '' Se existe SISCOTA, carregar OBSs, LOGs e Anexos
    '    If (Not pobjstcSis Is Nothing) Then
    '        '' Carrega OBSs
    '        Dim objObsA As clsA46V750A
    '        Set objObsA = New clsA46V750A
    '        Set objObsA.M46V999 = Me.M46V999
    '
    '        Dim objLogA As clsA46V751A
    '        Set objLogA = New clsA46V751A
    '        Set objLogA.M46V999 = Me.M46V999
    '
    '        lsSQl = objObsA.mfFormatarConsultaSQLByIdSiscota(pobjstcSis.IdSiscota)
    '
    '        If M46V999.gfObtRegistro(ConexaoBD, lsSQl, lrstObs, UltimoErro) = False Then GoTo Fim
    '
    '        Do While Not lrstObs.EOF
    '            Set objObs = New stcA46V750B
    '            Call objObsA.mpCarregarAtributos(objObs, lrstObs)
    '
    '            '' ------------------------------------------------------
    '            '' Carrega LOGs
    '            'Dim objLogA As clsA46V751A
    '            'Set objLogA = New clsA46V751A
    '            'Set objLogA.M46V999 = Me.M46V999
    '            lsSQl = objLogA.mfFormatarConsultaSQL(objObs.IdObs)
    '
    '            If M46V999.gfObtRegistro(ConexaoBD, lsSQl, lrstLog, UltimoErro) = False Then GoTo Fim
    '
    '            Do While Not lrstLog.EOF
    '                Set objLog = New stcA46V751B
    '                Call objLogA.mpCarregarAtributos(objLog, lrstLog)
    '
    '                objObs.ColecaoLog.Add objLog
    '
    '                lrstLog.MoveNext
    '            Loop
    '
    '            Set lrstLog = Nothing
    '            '' Fim LOGs
    '            '' ------------------------------------------------------
    '
    '            '' Carrega Anexos
    '            Dim objAnexoA As clsA46V752A
    '            Set objAnexoA = New clsA46V752A
    '            Set objAnexoA.M46V999 = Me.M46V999
    '
    '            lsSQl = objAnexoA.mfFormatarConsultaSQL(objObs.IdObs)
    '
    '            If M46V999.gfObtRegistro(ConexaoBD, lsSQl, lrstAnexo, UltimoErro) = False Then GoTo Fim
    '
    '            Do While Not lrstAnexo.EOF
    '                Set objAnexo = New stcA46V752B
    '                Call objAnexoA.mpCarregarAtributos(objAnexo, lrstAnexo)
    '
    '                objObs.ColecaoAnexo.Add objAnexo
    '
    '                lrstAnexo.MoveNext
    '            Loop
    '
    '            Set lrstAnexo = Nothing
    '            '' Fim Anexos
    '
    '            pobjstcSis.ColecaoObs.Add objObs
    '
    '            lrstObs.MoveNext
    '        Loop
    '
    '        Set lrstObs = Nothing
    '    End If
    '
Fim:
    '' Fim com sucesso
    mfCarregarMinimo = True
    Exit Function

ControleErros:
    Set lrstObs = Nothing
    mfCarregarMinimo = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V750B.mfCarregarMinimo ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V750B.mfCarregarMinimo ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
End Function

Public Function mfCarregar(ByRef pobjstcSis As stcA46V753B, ByRef Mensagem As String) As Boolean
    Dim lrstSis As clsYasRecordSet      'Registro.
    Dim lrstObs As clsYasRecordSet      'Registro.
    Dim lrstLog As clsYasRecordSet      'Registro.
    Dim lrstAnexo As clsYasRecordSet    'Registro.

    Dim objObs As stcA46V750B
    Dim objLog As stcA46V751B
    Dim objAnexo As stcA46V752B
    Dim lsSQl As String                 'SQL.

    On Error GoTo ControleErros

    If Not mfConectar Then GoTo ControleErros

    lsSQl = mfFormatarConsultaSQL(pobjstcSis)

    If M46V999.gfObtRegistro(ConexaoBD, lsSQl, lrstSis, UltimoErro) = False Then GoTo ControleErros

    If Not lrstSis.EOF Then
        Call mpCarregarAtributos(pobjstcSis, lrstSis)
    Else
        Set lrstSis = Nothing
        Mensagem = "Registro não localizado com Id=" & M46V999.gfTratarNulo(pobjstcSis.IdSiscota, tpoNumerico)
        GoTo Fim
    End If

    '' Se existe SISCOTA, carregar OBSs, LOGs e Anexos
    If (Not pobjstcSis Is Nothing) Then
        '' Carrega OBSs
        Dim objObsA As clsA46V750A
        Set objObsA = New clsA46V750A
        Set objObsA.M46V999 = Me.M46V999
        
        Dim objLogA As clsA46V751A
        Set objLogA = New clsA46V751A
        Set objLogA.M46V999 = Me.M46V999
        
        lsSQl = objObsA.mfFormatarConsultaSQLByIdSiscota(pobjstcSis.IdSiscota)
        
        If M46V999.gfObtRegistro(ConexaoBD, lsSQl, lrstObs, UltimoErro) = False Then GoTo Fim

        Do While Not lrstObs.EOF
            Set objObs = New stcA46V750B
            Call objObsA.mpCarregarAtributos(objObs, lrstObs)

            '' ------------------------------------------------------
            '' Carrega LOGs
            'Dim objLogA As clsA46V751A
            'Set objLogA = New clsA46V751A
            'Set objLogA.M46V999 = Me.M46V999
            lsSQl = objLogA.mfFormatarConsultaSQL(objObs.IdObs)

            If M46V999.gfObtRegistro(ConexaoBD, lsSQl, lrstLog, UltimoErro) = False Then GoTo Fim

            Do While Not lrstLog.EOF
                Set objLog = New stcA46V751B
                Call objLogA.mpCarregarAtributos(objLog, lrstLog)

                objObs.ColecaoLog.Add objLog
                
                lrstLog.MoveNext
            Loop
            
            Set lrstLog = Nothing
            '' Fim LOGs
            '' ------------------------------------------------------
            
            '' Carrega Anexos
            Dim objAnexoA As clsA46V752A
            Set objAnexoA = New clsA46V752A
            Set objAnexoA.M46V999 = Me.M46V999
            
            lsSQl = objAnexoA.mfFormatarConsultaSQL(objObs.IdObs)
            
            If M46V999.gfObtRegistro(ConexaoBD, lsSQl, lrstAnexo, UltimoErro) = False Then GoTo Fim
            
            Do While Not lrstAnexo.EOF
                Set objAnexo = New stcA46V752B
                Call objAnexoA.mpCarregarAtributos(objAnexo, lrstAnexo)
                            
                objObs.ColecaoAnexo.Add objAnexo
            
                lrstAnexo.MoveNext
            Loop
            
            Set lrstAnexo = Nothing
            '' Fim Anexos
            
            pobjstcSis.ColecaoObs.Add objObs
                        
            lrstObs.MoveNext
        Loop
        
        Set lrstObs = Nothing
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

Public Function mfCarregarObs(ByRef pobjstcSis As stcA46V753B, ByRef Mensagem As String) As Boolean
    Dim lrstSis As clsYasRecordSet      'Registro.
    Dim lrstObs As clsYasRecordSet      'Registro.
    Dim lrstLog As clsYasRecordSet      'Registro.
    Dim lrstAnexo As clsYasRecordSet    'Registro.
    
    Dim objObs As stcA46V750B
    Dim objLog As stcA46V751B
    Dim objAnexo As stcA46V752B
    Dim lsSQl As String                 'SQL.
    Dim i As Integer

    On Error GoTo ControleErros
    
    If Not mfConectar Then GoTo ControleErros
    
    lsSQl = mfFormatarConsultaSQL(pobjstcSis)
    
    If M46V999.gfObtRegistro(ConexaoBD, lsSQl, lrstSis, UltimoErro) = False Then GoTo ControleErros
    
    If Not lrstSis.EOF Then
        Call mpCarregarAtributos(pobjstcSis, lrstSis)
    Else
        Set lrstSis = Nothing
        Mensagem = "Registro não localizado com Id=" & M46V999.gfTratarNulo(pobjstcSis.IdSiscota, tpoNumerico)
        GoTo Fim
    End If
    
    '' Se existe SISCOTA, carregar OBSs, LOGs e Anexos
    If (Not pobjstcSis Is Nothing) Then
        '' Carrega OBSs
        Dim objObsA As clsA46V750A
        Set objObsA = New clsA46V750A
        Set objObsA.M46V999 = Me.M46V999
        
        Dim objLogA As clsA46V751A
        Set objLogA = New clsA46V751A
        Set objLogA.M46V999 = Me.M46V999
        
        lsSQl = objObsA.mfFormatarConsultaSQL(pobjstcSis.ColecaoObs(1))
        
        If M46V999.gfObtRegistro(ConexaoBD, lsSQl, lrstObs, UltimoErro) = False Then GoTo Fim
        
        Do While Not lrstObs.EOF
            Set objObs = New stcA46V750B
            Call objObsA.mpCarregarAtributos(objObs, lrstObs)
            
            '' ------------------------------------------------------
            '' Carrega LOGs
            'Dim objLogA As clsA46V751A
            'Set objLogA = New clsA46V751A
            'Set objLogA.M46V999 = Me.M46V999
            lsSQl = objLogA.mfFormatarConsultaSQL(objObs.IdObs)
            
            If M46V999.gfObtRegistro(ConexaoBD, lsSQl, lrstLog, UltimoErro) = False Then GoTo Fim
            
            Do While Not lrstLog.EOF
                Set objLog = New stcA46V751B
                Call objLogA.mpCarregarAtributos(objLog, lrstLog)
                   
                objObs.ColecaoLog.Add objLog
                
                lrstLog.MoveNext
            Loop
            
            Set lrstLog = Nothing
            '' Fim LOGs
            '' ------------------------------------------------------
            
            '' Carrega Anexos
            Dim objAnexoA As clsA46V752A
            Set objAnexoA = New clsA46V752A
            Set objAnexoA.M46V999 = Me.M46V999
            
            lsSQl = objAnexoA.mfFormatarConsultaSQL(objObs.IdObs)
            
            If M46V999.gfObtRegistro(ConexaoBD, lsSQl, lrstAnexo, UltimoErro) = False Then GoTo Fim
            
            Do While Not lrstAnexo.EOF
                Set objAnexo = New stcA46V752B
                Call objAnexoA.mpCarregarAtributos(objAnexo, lrstAnexo)
                            
                objObs.ColecaoAnexo.Add objAnexo
            
                lrstAnexo.MoveNext
            Loop
            
            Set lrstAnexo = Nothing
            '' Fim Anexos
            
            '' remover obs da chamada
            pobjstcSis.ColecaoObs.Remove (1)
                        
            pobjstcSis.ColecaoObs.Add objObs
                        
            lrstObs.MoveNext
        Loop
        
        Set lrstObs = Nothing
    End If
    
Fim:
    '' Fim com sucesso
    mfCarregarObs = True
    Exit Function

ControleErros:
    Set lrstObs = Nothing
    mfCarregarObs = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V750B.mfCarregarObs ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V750B.mfCarregarObs ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
End Function


Public Function mfSalvar(ByRef pobjstcSis As stcA46V753B, _
                         ByVal pLogin As String, _
                         ByRef Mensagem As String) As Boolean

    Dim lsSQl As String           'Query.
    
    Mensagem = ""
    
    On Local Error GoTo ControleErros
    
    lsSQl = "mfConectar"
    If Not mfConectar Then GoTo ControleErros
    
    If pobjstcSis.IdSiscota > 0 Then
        pobjstcSis.IdSiscota = pobjstcSis.IdSiscota
        
        If (pobjstcSis.ColecaoObs(1).IdObs > 0) Then
            lsSQl = mfFormatarAtualizacaoSQLObs(pobjstcSis.ColecaoObs(1), pLogin)
        Else
            lsSQl = mfFormatarInclusaoSQLObs(pobjstcSis.ColecaoObs(1), pLogin)
        End If
    Else
        lsSQl = mfFormatarInclusaoSQL(pobjstcSis, pLogin)
    End If
    
    If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then
        GoTo ControleErros
    End If
    
    mfSalvar = True
    Exit Function
    
ControleErros:
    Mensagem = "A46V753A.mfSalvar ERRO - " & UltimoErro & vbCrLf & " SQL: " & lsSQl
    Call M46V999.gpGraLog(0, Mensagem)

End Function


'Public Function mfFormatarExclusaoSQL(ByRef pobjstcSis As stcA46V753B, _
'                                       ByVal pLogin As String) As String
'    Dim lDelete As String 'Query
'
'    If pobjstcSis.IdSiscota > 0 Then
'        '' Remocao Logica: status=2
'        lDelete = "UPDATE RamosDiversos..Tab_Ped_Obs_SISCOTA SET "
'        lDelete = lDelete & "Status = 2 where IdObs = " & M46V999.gfFormatarCamposSQL(pobjstc.IdObs) & vbCrLf
'
'        '' O correto seria colocar na classe 751A, mas ficará aqui para simplificar
'        lDelete = lDelete & "INSERT INTO RamosDiversos..Tab_Ped_Obs_SISCOTA_LOG "
'        lDelete = lDelete & "(IdSiscota, DataObs, TipoOperacao, Descricao, LoginObs) VALUES ("
'        lDelete = lDelete & M46V999.gfFormatarCamposSQL(pobjstcSis.IdSiscota) & ", getdate(), "
'        lDelete = lDelete & M46V999.gfFormatarCamposSQL(pobjstcSis.Status) & ", 'SISCOTA Removido', " & _
'                            M46V999.gfFormatarCamposSQL(pLogin) & ")" & vbCrLf
'    End If
'
'    mfFormatarExclusaoSQL = lDelete
'End Function

'' Nao havera exclusao de Tab_Ped_SISCOTA
'Public Function mfExcluir(ByRef pobjstc As stcA46V753B, _
'                          ByVal pLogin As String, _
'                          ByRef Mensagem As String) As Boolean
'
'    Dim lsSQl As String           'Query.
'
'    Mensagem = ""
'
'    On Local Error GoTo ControleErros
'
'    lsSQl = "mfConectar"
'    If Not mfConectar Then GoTo ControleErros
'
'    lsSQl = mfFormatarExclusaoSQL(pobjstc, pLogin)
'
'    If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then
'        GoTo ControleErros
'    End If
'
'    mfExcluir = True
'    Exit Function
'
'ControleErros:
'    Mensagem = "A46V750A.mfIncluir ERRO - " & UltimoErro & vbCrLf & " SQL: " & lsSQl
'    Call M46V999.gpGraLog(0, Mensagem)
'
'End Function

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



'' Poderiamos fazer uma classe especifica para a carga, mas manteremos aqui para evitar voltas desnecessarias
Public Function mfCarregarObservacoes(ByRef objLog As Collection, ByVal nossoNumeroStr As String, ByRef Mensagem As String) As Boolean
    Dim lrstObs As clsYasRecordSet      'Registro.
    Dim lsSQl As String                 'SQL.
    Dim obj As stcA46V750Brpt
    Dim lRegAfetados As Long
    lRegAfetados = 0

    On Error GoTo ControleErros
    
    'Set M46V999 = gM46V999
    
    If Not mfConectar Then GoTo ControleErros
    
    lsSQl = "exec RamosDiversos..proc_BuscarObservacoesSiscota '" & nossoNumeroStr & "', 1"
    
    If M46V999.gfObtRegistro(ConexaoBD, lsSQl, lrstObs, UltimoErro) = False Then GoTo ControleErros
    
    If lrstObs.EOF Then
        Set lrstObs = Nothing
        Mensagem = "Registro não localizado com NossoNúmero = " & M46V999.gfTratarNulo(nossoNumeroStr, tpoNumerico)
    Else
        Do While Not lrstObs.EOF
            Set obj = New stcA46V750Brpt
            Call mpCarregarAtributosRpt(obj, lrstObs)
                        
            objLog.Add obj
            
            lrstObs.MoveNext
        Loop
        
        Set lrstObs = Nothing
    End If
    
Fim:
    '' Fim com sucesso
    mfCarregarObservacoes = True
    Exit Function

ControleErros:
    'If Not lrstObs = Nothing Then
        Set lrstObs = Nothing
    'End If
    
    mfCarregarObservacoes = False
    
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V750Brpt.mfCarregarObservacoes ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V750Brpt.mfCarregarObservacoes ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
End Function

Public Function mfAtualizaStatusSiscota(ByVal pobjstc As stcA46V753B, _
                                        ByVal pLogin As String) As Boolean
    
    Dim lsSQl As String
    
    On Error GoTo ControleErros
    
    mfAtualizaStatusSiscota = False
    
    If Not mfConectar Then GoTo ControleErros
    
    lsSQl = mfFormatarAtualizacaoSQLSis(pobjstc, pLogin)
    
    If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then
        GoTo ControleErros
    Else
        lsSQl = mfFormatarAtualizacaoCodOrigemObsSQL(1, pobjstc.ColecaoObs(1).IdObs, pLogin)
        
        If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then
            GoTo ControleErros
        Else
            mfAtualizaStatusSiscota = True
            Exit Function
        End If
    End If

ControleErros:
    
    mfAtualizaStatusSiscota = False
    
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "mfAtualizaStatusSiscota ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "mfAtualizaStatusSiscota ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
    
End Function

Public Sub mpCarregarAtributosRpt(ByRef pobjstc As stcA46V750Brpt, prst As clsYasRecordSet)
    pobjstc.IdSiscota = M46V999.gfTratarNulo(prst("IdSiscota"), tpoNumerico)
    pobjstc.IdObs = M46V999.gfTratarNulo(prst("IdObs"), tpoNumerico)
    pobjstc.IdGrupo = M46V999.gfTratarNulo(prst("IdGrupo"), tpoNumerico)
    pobjstc.nosso_numero = M46V999.gfTratarNulo(prst("Nosso_Numero"), tpoString)
    pobjstc.Num_Ped = M46V999.gfTratarNulo(prst("Num_Ped"), tpoNumerico)
    pobjstc.Siscota = M46V999.gfTratarNulo(prst("SISCOTA"), tpoString)
    pobjstc.STATUS = M46V999.gfTratarNulo(prst("Status"), tpoNumerico)
    pobjstc.CodOrigem = M46V999.gfTratarNulo(prst("CodOrigem"), tpoNumerico)
    pobjstc.Observacao = M46V999.gfTratarNulo(prst("Observacao"), tpoString)
    pobjstc.Data = M46V999.gfTratarNulo(prst("Data"), TpoData)
    pobjstc.Login = M46V999.gfTratarNulo(prst("Login"), tpoString)
End Sub
