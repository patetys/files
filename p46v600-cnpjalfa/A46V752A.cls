VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA46V752A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: clsA46V752A - Tab_Ped_Obs_SISCOTA_Anexo
' Data/Hora: 19/10/2017
' Sistema: Multi
' Juliano Iwai
' ======================================================================================================================================================

Public ConexaoBD As ADODB.Connection    ' Conexão com o Banco de Dados
Public Colecao As Collection            ' Coleção de objetos do tipo stcA46V749B
Public UltimoErro As String             ' Descrição do último erro gerado

Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "A46V752A"
End Property

Public Function mfFormatarConsultaSQL(ByVal IdObs As Long) As String
    Dim lSelect As String 'Query.
    lSelect = "SELECT IdAnexo, IdObs, DataAnexo, NomeArquivo, PastaArquivo, NomeArquivoFisico FROM RamosDiversos..Tab_Ped_SISCOTA_Anexo(nolock)"
    lSelect = lSelect & " Where IdObs = " & M46V999.gfFormatarCamposSQL(IdObs)
    lSelect = lSelect & " Order by IdAnexo DESC "

    mfFormatarConsultaSQL = lSelect
End Function

Public Function mfFormatarConsultaSQLNossoNumero(ByVal strNossoNumero As String) As String
    Dim lSelect As String 'Query.
    lSelect = "select top 50 Ane.IdAnexo, Ane.IdObs, Ane.DataAnexo, Ane.NomeArquivo, Ane.PastaArquivo, Ane.NomeArquivoFisico " & vbCrLf
    lSelect = lSelect & " Sis.Siscota , row_number() over(order by idAnexo) as [SEQUENCIA] " & vbCrLf
    lSelect = lSelect & "from RamosDiversos..Tab_Ped_Siscota_Anexo Ane (nolock) " & vbCrLf
    lSelect = lSelect & " inner join RamosDiversos..Tab_Ped_Siscota_Obs Obs (nolock) on Ane.IdObs = Obs.IdObs " & vbCrLf
    lSelect = lSelect & " inner join RamosDiversos..Tab_Ped_Siscota Sis (nolock) on Obs.IdSiscota = Sis.IdSiscota " & vbCrLf
    lSelect = lSelect & "where Sis.Nosso_Numero='" & strNossoNumero & "'"

    mfFormatarConsultaSQLNossoNumero = lSelect
End Function

'Public Function mfFormatarInclusaoSQL(ByRef pobjstc As stcA46V750B, _
'                                       ByVal pLogin As String) As String
'    Dim lInsert As String 'Query.
'    Dim pobjstcAnexo As stcA46V752B
'
'    If (pobjstc.ColecaoAnexo.Count > 0) Then
'    '' Remove todos e reinsere.
'    lUpdate = lUpdate & "DELETE FROM RamosDiversos..Tab_Ped_SISCOTA_ANEXO where IdObs=" & M46V999.gfFormatarCamposSQL(pobjstc.IdObs) & vbCrLf
'
'    For Each pobjstcAnexo In pobjstc.ColecaoAnexo
'        lUpdate = lUpdate & "INSERT INTO RamosDiversos..Tab_Ped_SISCOTA_ANEXO "
'        lUpdate = lUpdate & "(IdObs, DataAnexo, NomeArquivo, PastaArquivo) VALUES ("
'        lUpdate = lUpdate & M46V999.gfFormatarCamposSQL(pobjstcAnexo.IdObs) & ", "
'        lUpdate = lUpdate & M46V999.gfFormatarCamposSQL(pobjstcAnexo.DataAnexo) & ", "
'        lUpdate = lUpdate & M46V999.gfFormatarCamposSQL(pobjstcAnexo.NomeArquivo) & ", "
'        lUpdate = lUpdate & M46V999.gfFormatarCamposSQL(pobjstcAnexo.PastaArquivo) & ") " & vbCrLf
'    Next
'    End If
'
'    mfFormatarInclusaoSQL = lInsert
'End Function

Public Sub mpCarregarAtributos(ByRef pobjstc As stcA46V752B, prst As clsYasRecordSet)
    pobjstc.IdAnexo = M46V999.gfTratarNulo(prst("IdAnexo"), tpoNumerico)
    pobjstc.IdObs = M46V999.gfTratarNulo(prst("IdObs"), tpoNumerico)
    pobjstc.DataAnexo = M46V999.gfTratarNulo(prst("DataAnexo"), TpoData)
    pobjstc.NomeArquivo = M46V999.gfTratarNulo(prst("NomeArquivo"), tpoString)
    pobjstc.NomeArquivoFisico = M46V999.gfTratarNulo(prst("NomeArquivoFisico"), tpoString)
    pobjstc.PastaArquivo = M46V999.gfTratarNulo(prst("PastaArquivo"), tpoString)
    If Not prst("Sequencia") Is Nothing Then
        pobjstc.Sequencia = M46V999.gfTratarNulo(prst("Sequencia"), tpoString)
    End If
End Sub

Public Function mfCarregarColecao(ByRef pobjstc As stcA46V752B) As Boolean
    Dim lobjstc As stcA46V752B      'Objeto da classe de estrutura stcA46V751B.
    Dim lrst As clsYasRecordSet     'Registro.
    Dim lsSQl As String             'SQL.

    On Error GoTo ControleErros
    
    If Not mfConectar Then GoTo ControleErros
    
    lsSQl = mfFormatarConsultaSQL(pobjstc.IdObs)
    If M46V999.gfObtRegistro(ConexaoBD, lsSQl, lrst, UltimoErro) = False Then GoTo ControleErros
    
    Do While Not lrst.EOF
        Set lobjstc = New stcA46V752B
        
        Call mpCarregarAtributos(lobjstc, lrst)
        Colecao.Add Item:=lobjstc
        lrst.MoveNext
    Loop
    
    Set lobjstc = Nothing
    Set lrst = Nothing
    mfCarregarColecao = True
    Exit Function
    
ControleErros:
    Set lobjstc = Nothing
    Set lrst = Nothing
    mfCarregarColecao = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V752A.mfCarregarColecao ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V752A.mfCarregarColecao ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
End Function

Public Function mfCarregarColecaoByNossoNumero(ByVal strNossoNumero As String) As Boolean
    Dim lobjstc As stcA46V752B      'Objeto da classe de estrutura stcA46V751B.
    Dim lrst As clsYasRecordSet     'Registro.
    Dim lsSQl As String             'SQL.

    On Error GoTo ControleErros
    
    If Not mfConectar Then GoTo ControleErros
    
    lsSQl = mfFormatarConsultaSQLNossoNumero(strNossoNumero)
    If M46V999.gfObtRegistro(ConexaoBD, lsSQl, lrst, UltimoErro) = False Then GoTo ControleErros
    
    Do While Not lrst.EOF
        Set lobjstc = New stcA46V752B
        
        Call mpCarregarAtributos(lobjstc, lrst)
        Colecao.Add Item:=lobjstc
        lrst.MoveNext
    Loop
    
    Set lobjstc = Nothing
    Set lrst = Nothing
    mfCarregarColecaoByNossoNumero = True
    Exit Function
    
ControleErros:
    Set lobjstc = Nothing
    Set lrst = Nothing
    mfCarregarColecaoByNossoNumero = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V752A.mfCarregarColecao ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V752A.mfCarregarColecao ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
End Function

Private Sub Class_Initialize()
    Set Colecao = New Collection
    Set mClsM46V999 = Nothing
End Sub

Private Sub Class_Terminate()
    Set Colecao = Nothing
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




