Attribute VB_Name = "basM46V995"
'Option Explicit
'' ======================================================================================================================================================
'' Modulo: basM46V995 --> Modulo para carregamento do DE / Para
'' Data/Hora: 05/08/2014
'' Sistema: SyasMulti Web
'' Rodrigo Berçot
'' ======================================================================================================================================================
'
'Public gRst_Tbm_De_Para As ADODB.Recordset
'Private mClsM46V999 As clsM46V999
'Public gCodTarifa As Integer
'Public gCodRamo As Integer
'Public gCodplano As Integer
'
'Public Property Set M46V999_Cnx(pClsM46V999 As clsM46V999)
'    Set mClsM46V999 = pClsM46V999
'End Property
'
'Public Property Get M46V999_Cnx() As clsM46V999
'
'    If mClsM46V999 Is Nothing Then
''        Err.Raise 91, Class_Na, "Objeto M46V999 não atribuído para a classe (" & Me.Class_Name & ")"
'        '***Esta Classe DEVE utilizar uma instancia da classe clsM46V999 já
'        '   iniciada no escopo onde esta classe foi instanciada.
'        'Atribua a o objeto (clsM46V999 ja instanciado) no momento que eu inicializou esta classe (atraves do Property SET)!!!
'    End If
'
'Retorno:
'    Set M46V999_Cnx = mClsM46V999
'End Property
'
''Sub mfCarregarDePara()
''    Dim lSel_De_Para As String
''    Dim lP00Multi_Tplano As ADODB.Connection
''    Dim lMensagem As String
''    Dim Rst_Tab_Ped As ADODB.Recordset
''
''    Set Rst_Tab_Ped = New ADODB.Recordset
''    Call M46V999_Cnx.gpGraLog(0, "Carga De->Para")
''
''    If M46V999_Cnx.gfAbrBasDados("P00MULTI", M46V999_Cnx.gTipP00MULTI, M46V999_Cnx.gEndP00MULTI, lP00Multi_Tplano, lMensagem) = False Then Call M46V999_Cnx.gpGraLog(2, lMensagem): Exit Sub
''
''    lSel_De_Para = "EXEC dbo.proc_Buscar_De_Para "
''    lSel_De_Para = lSel_De_Para & gCod_Ramo & ","
''    lSel_De_Para = lSel_De_Para & gCod_plano & ","
''    lSel_De_Para = lSel_De_Para & gCod_Tarifa & ","
''    lSel_De_Para = lSel_De_Para & "NULL,"
''    lSel_De_Para = lSel_De_Para & "NULL,"
''    lSel_De_Para = lSel_De_Para & "NULL"
''
''    On Error Resume Next
''    Err.Clear
''    Set gRst_Tbm_De_Para = New ADODB.Recordset
''    gRst_Tbm_De_Para.CursorLocation = adUseClient
''    Call gRst_Tbm_De_Para.Open(lSel_De_Para, lP00Multi_Tplano)    'Este recordSet esta usando a propriedade Filter (nao implementado no YasRecordSet)
''    Set gRst_Tbm_De_Para.ActiveConnection = Nothing
''    If Err <> 0 Then
''        Call M46V999_Cnx.gpGraLog(2, Err.Number & " - " & Err.Description): Exit Sub
''    End If
''
''    Set lP00Multi_Tplano = Nothing
''End Sub
'
''Public Function gfBuscarDePara(ByVal pRetorno As String, ByRef pTipoDePara As Integer, Optional ByVal pCompl As String, Optional pRetornaComplemento As Boolean) As String
''    'Função reponsavel por trazer o De Para correspondente ao codigo recebido por parâmetro.
''        'FILTRO DE/PARA
''        'TIPO_DE_PARA =  1 - Coberturas
''        '                2 - Classificações
''        '                3 - Planos
''        If gCod_Tarifa >= 1080 Then
''            If gRst_Tbm_De_Para Is Nothing Then
''                If pCompl <> "" Then
''                    gfBuscarDePara = pCompl
''                Else
''                    gfBuscarDePara = pRetorno
''                End If
''                Exit Function
''            End If
''            If gRst_Tbm_De_Para.RecordCount > 0 Then
''                gRst_Tbm_De_Para.MoveFirst
''                If pCompl = "" Then
''                    gRst_Tbm_De_Para.Filter = "DE = '" & pRetorno & "' AND TIPO_DE_PARA = " & pTipoDePara
''                    If gRst_Tbm_De_Para.RecordCount > 0 Then
''                        gfBuscarDePara = gRst_Tbm_De_Para!Para
''                    Else
''                        gfBuscarDePara = pRetorno
''                    End If
''                Else
''                    gRst_Tbm_De_Para.Filter = "DE = '" & pRetorno & "' AND DE_COMPL = '" & pCompl & "' AND TIPO_DE_PARA = " & pTipoDePara
''
''                    If gRst_Tbm_De_Para.RecordCount > 0 Then
''                        If pRetornaComplemento = False Then
''                            gfBuscarDePara = gRst_Tbm_De_Para!Para
''                        Else
''                            gfBuscarDePara = gRst_Tbm_De_Para!Para_Compl
''                        End If
''                    Else
''                        If pRetornaComplemento Then
''                            gfBuscarDePara = pCompl
''                        Else
''                            gfBuscarDePara = pRetorno
''                        End If
''                    End If
''
''                End If
''
''                gRst_Tbm_De_Para.Filter = adFilterNone
''            Else
''                If pCompl <> "" Then
''                    gfBuscarDePara = pCompl
''                Else
''                    gfBuscarDePara = pRetorno
''                End If
''
''            End If
''        Else
''            If pCompl <> "" Then
''                gfBuscarDePara = pCompl
''            Else
''                gfBuscarDePara = pRetorno
''            End If
''        End If
''
''End Function
