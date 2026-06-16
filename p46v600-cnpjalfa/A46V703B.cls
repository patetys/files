VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V703B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================

' Class Name: stcA46V703B --> Tabela Tab_Ped_Clau
' Data/Hora: 31/08/2005 - 11:51:43
' Sistema: Multi
' Thais Juren Gomes

' ======================================================================================================================================================


' ======================================================================================================================================================
Public nosso_numero As String           'Nosso número.
Public Cod_Clau As Long         'Código da cláusula.
Public Num_Ped As Long        'Número do pedido.
Public Num_Apol As Double         'Sem uso - Número da apólice.
Public Num_Endo As Long         'Sem uso - Número do endosso.
Public Num_Item As Long         'Número do item.
'Public NUM_PI As Long       'Número do PI.
Public NUM_PI As Double       'Número do PI.
Public Cod_Ramo As e_CodRamo    'Código do ramo.
Public Cod_Org As Integer     'Código do orgão.
Public Descricao As String
''Public UltimoErro       As String

' ======================================================================================================================================================
Private ObjStcPai As stcA46V708B
' ======================================================================================================================================================

Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "A46V703B"
End Property

Public Property Get Class_ID() As String

    Class_ID = CStr(nosso_numero) & _
               CStr(Format(Num_Item, "00000")) & _
               CStr(Format(Cod_Clau, "00000"))
End Property

Public Property Get gCopiaClau(pNumItem As Long) As stcA46V703B
    Dim lobjstcClau As stcA46V703B
    Set lobjstcClau = New stcA46V703B
    Set lobjstcClau.M46V999 = M46V999

    With lobjstcClau
        .Cod_Clau = Cod_Clau
        .Cod_Org = Cod_Org
        .Cod_Ramo = Cod_Ramo
        .nosso_numero = nosso_numero
        .Num_Apol = Num_Apol
        .Num_Endo = Num_Endo
        .Num_Item = pNumItem
        .Num_Ped = Num_Ped
        .NUM_PI = NUM_PI
    End With
    Set gCopiaClau = lobjstcClau
End Property

Public Property Get gDescClausula() As String
    Dim lbdTab_Clau As ADODB.Connection    'Arquivo: P00Tarifa.
    Dim lMensagem As String         'Mensagem.
    Dim lrsTab_Clau As clsYasRecordSet  'Registro : TAB_CLAU
    Dim lSelect As String

    If Not M46V999.gfPreenchido(Descricao) Then
        If Cod_Clau <> 0 Then
            If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, lbdTab_Clau, lMensagem) = False Then
                Call M46V999.gpGraLog(2, lMensagem)
                Exit Property
            End If
            lSelect = "Select * from Tab_Clau WITH (NOLOCK) where cod_ramo = " & Me.Cod_Ramo
            lSelect = lSelect & " and cod_clau = " & Me.Cod_Clau
            If Not M46V999.gfObtRegistro(lbdTab_Clau, lSelect, lrsTab_Clau, lMensagem) Then
                Call M46V999.gpGraLog(1, lMensagem)
                Exit Property
            End If
            If Not lrsTab_Clau.EOF Then
                Descricao = lrsTab_Clau("DSC_CLau")
            End If
        End If
    End If
    gDescClausula = Descricao
End Property

Public Property Get gEhAss24h() As Boolean
    Dim lbdAss24hrs As ADODB.Connection    'Arquivo: P00Tarifa.
    Dim lMensagem As String         'Mensagem.
    Dim lrsAss24hrs As clsYasRecordSet  'Registro : TAB_CLAU
    Dim lSelect As String
    Dim lretorno As Boolean

    lretorno = False
    
'    If mClsM46V999 Is Nothing Then
'        Set M46V999 = M46V999_Cnx
'    End If
        
    If Not mClsM46V999 Is Nothing And Cod_Clau <> 0 Then
        If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, lbdAss24hrs, lMensagem) = False Then
            Call M46V999.gpGraLog(2, lMensagem)
            Exit Property
        End If
        lSelect = "EXEC SP_RetornarAssistencias24Horas "
        lSelect = lSelect & Me.M46V999.gCodProduto & ","
        lSelect = lSelect & 0 & ","
        lSelect = lSelect & Me.Cod_Ramo & ","
        lSelect = lSelect & Me.M46V999.M46V111.gCod_plano
        
        If Not M46V999.gfObtRegistro(lbdAss24hrs, lSelect, lrsAss24hrs, lMensagem) Then
            Call M46V999.gpGraLog(1, lMensagem)
            Exit Property
        End If
        
        Do Until lrsAss24hrs.EOF
            If lrsAss24hrs!CodigoClausula = Me.Cod_Clau Then
                lretorno = True
                Exit Do
            End If
            
            lrsAss24hrs.MoveNext
        Loop
    End If

    gEhAss24h = lretorno
End Property

Public Property Set ObjPai(pobjstc As stcA46V708B)
    Set ObjStcPai = pobjstc
End Property

Public Property Get ObjPai() As stcA46V708B
    Set ObjPai = ObjStcPai
End Property

Public Property Get Cod_Unicidade() As Long
    If Not ObjStcPai Is Nothing Then
        Cod_Unicidade = ObjStcPai.Cod_Unicidade
    End If
End Property

Public Function getRpt() As stcA46V703Brpt
    Static lrpt As stcA46V703Brpt

    If lrpt Is Nothing Then
        Set lrpt = New stcA46V703Brpt
        Set lrpt.stcOrigem = Me
    End If

    Set getRpt = lrpt

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

Private Sub Class_Initialize()
    Set mClsM46V999 = Nothing
End Sub

Private Sub Class_Terminate()
    Set mClsM46V999 = Nothing
End Sub
