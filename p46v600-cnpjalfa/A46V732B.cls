VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V732B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================

' Class Name: stcA46V732B --> Tabela Tab_Ped_LMI
' Data/Hora: 09/11/2012 - 16:35
' Sistema: Multi
' Thais Juren Gomes Novaes

' ======================================================================================================================================================


' ======================================================================================================================================================
Public nosso_numero As String       'Nosso número.
Public Cod_Org As Integer    'Código do órgão.
Public Cod_Ramo As e_CodRamo    'Código do ramo.
Public Num_Ped As Long    'Número do pedido.
Public Num_Apol As Double     'Sem uso - número da apólice.
Public Num_Endo As Long     'Sem uso - número do endosso.
Public Tab_Cobert As Integer
Public Cod_Cobert As Integer    'Código da cobertura.
Public Val_IS As Double    'Valor da importância segurada.
Public Per_Franq As Double    'Percentual da franquia.
Public Val_Franq_Min As Double        'Valor da franquia mínima.
Public Val_Franq_Max As Double        'Valor da franquia máxima.
Public Num_Meses_PI As Integer      'Número de Meses do Período Indenitário.
'Public NUM_PI As Long   'Número do PI.
Public NUM_PI As Double   'Número do PI.
Public TextoFranquia As String
' ======================================================================================================================================================
Private ObjStcPai As stcA46V702B
' ======================================================================================================================================================

Private mClsM46V999 As clsM46V999
Public Nom_Cobertura As String        'Nome da cobertura.
Public Nom_Franquia As String       'Nome da cobertura.
Public Tem_ As String      'Nome da cobertura.


Public Property Get Class_Name() As String
    Class_Name = "A46V732B"
End Property

'Private Function BuscaDadosAdicionais(ByRef pbdTAB_Cobert As ADODB.Connection, ByRef pMensagem As String) As Boolean
'    Dim lSelect As String
'    BuscaDadosAdicionais = False
'
'    On Error GoTo TrataErro
'
'    If M46V999.gPrograma = "D046O003" Then
'        M46V111.gCod_Regiao = 0
'        M46V111.gCod_plano = Me.ObjPai.Cod_Estip
'        M46V999.gCodProduto = Me.ObjPai.Cod_Produto
'    End If
'
'    lSelect = " Select distinct Cod_grupo_cobert, DSC_Cobert, Cod_Cobert, Per_Lim_Max_Cobert, COD_LIM_COBERT, TAB_COBERT  , PER_FRANQ , VAL_FRANQ_MIN from TAB_COBERT WITH (NOLOCK)"
'    lSelect = lSelect & " Where cod_Ramo = " & Cod_Ramo
'    lSelect = lSelect & " AND Cod_Cobert = " & Cod_Cobert
'    'If Val(Right(frmT46V101A.txtRisRegiao.Tag, 4)) <> 0 Then
'    If M46V111.gCod_Regiao <> 0 Then
'        lSelect = lSelect & " and tab_cobert = " & M46V111.gCod_Regiao
'    Else
'        'lSelect = lSelect & " and tab_cobert = " & frmT46V101A.cboPlano.ItemData(frmT46V101A.cboPlano.ListIndex)
'        If M46V111.gCod_plano = 0 Then
'            lSelect = lSelect & " and tab_cobert = " & Tab_Cobert
'        Else
'            lSelect = lSelect & " and tab_cobert = " & M46V111.gCod_plano
'        End If
'    End If
'    lSelect = lSelect & " and COD_TARIF_INI <= " & M46V999.gCodProduto
'    lSelect = lSelect & " and COD_TARIF_TER >= " & M46V999.gCodProduto
'
'    M46V999.gRst_Tab_Cobert.CursorLocation = adUseClient
'    Call gRst_Tab_Cobert.Open(lSelect, pbdTAB_Cobert)    'Este recordSet esta usando a propriedade Filter (nao implementado no YasRecordSet)
'    Set M46V999.gRst_Tab_Cobert.ActiveConnection = Nothing
'
''    If M46V999.gfObtRegistro(pbdTAB_Cobert, lSelect, prsTab_Cobert, pMensagem) = False Then
''        Call M46V999.gpFecha2(pbdTAB_Cobert)
''        Call M46V999.gpGraLog(2, pMensagem)
''        Screen.MousePointer = vbDefault
''        Exit Function
''    End If
'
'    BuscaDadosAdicionais = True
'
'TrataErro:
'        Call M46V999.gpFecha2(pbdTAB_Cobert)
'        Call M46V999.gpGraLog(2, Err.Description)
'        Screen.MousePointer = vbDefault
'        Exit Function
'
'End Function

' ======================================================================================================================================================

Public Property Get Class_ID() As String
    Class_ID = CStr(nosso_numero) & _
               CStr(Format(Cod_Cobert, "0000")) & _
               CStr(Format(Tab_Cobert, "00000"))
End Property

Public Property Get gNomeCobertura() As String
    Dim lbdTAB_Cobert As ADODB.Connection   'Arquivo: P00Tarifa.
    Dim lMensagem As String         'Mensagem.
    Dim lrsTab_Cobert As clsYasRecordSet    'Registro : TAB_COBERT.
    Dim lIndBasDados As String

    Static lsNomeCobertura As String

    If (lsNomeCobertura <> Nom_Cobertura) Or (Not M46V999.gfPreenchido(Nom_Cobertura)) Then
        If Cod_Cobert <> 0 Then
            If UCase(App.EXEName) = "D46O003" Then
                lIndBasDados = "P00MPED"
            Else
                lIndBasDados = "P00MULTI"
            End If
            If M46V999.gfAbrBasDados(lIndBasDados, M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, lbdTAB_Cobert, lMensagem) = False Then
                Call M46V999.gpGraLog(2, lMensagem)
                Exit Sub
            End If
            
            If M46V999.gRst_Tab_Cobert.State = 0 Then
                Call M46V999.BuscaDadosAdicionais(lbdTAB_Cobert, lMensagem)
            ElseIf M46V999.gRst_Tab_Cobert.RecordCount = 0 Then
                Call M46V999.BuscaDadosAdicionais(lbdTAB_Cobert, lMensagem)
                End If
'            If BuscaDadosAdicionais(lbdTAB_Cobert, lrsTab_Cobert, lMensagem) Then
'                If Not lrsTab_Cobert.EOF Then
                 If Not M46V999.gRst_Tab_Cobert.EOF Then
                    M46V999.gRst_Tab_Cobert.Filter = "cod_cobert = " & Cod_Cobert
                    Nom_Cobertura = M46V999.gRst_Tab_Cobert("DSC_Cobert")
                    lsNomeCobertura = Nom_Cobertura    '!!! Guarda na variavel statica
            End If
'            End If
        End If
    End If

    gNomeCobertura = lsNomeCobertura

End Property


Public Property Get gNomeFranquia() As String

    Static lsNomeFranquia As String

    If (lsNomeFranquia <> Nom_Franquia) Or (Not M46V999.gfPreenchido(Nom_Franquia)) Then
    End If

    gNomeFranquia = lsNomeFranquia

End Property


Public Property Get gCopiaCobert(pNumItem As Long) As stcA46V704B
    Dim lobjstcCobert As stcA46V704B

    Set lobjstcCobert = New stcA46V704B
    Set lobjstcCobert.M46V999 = M46V999

    With lobjstcCobert
        .nosso_numero = nosso_numero
        .Cod_Cobert = Cod_Cobert
        .Num_Ped = .Num_Ped
        .Cod_Org = Cod_Org
        .Cod_Ramo = Cod_Ramo
        .Num_Apol = Num_Apol
        .Num_Endo = Num_Endo
        .Num_Item = pNumItem
        .Tab_Cobert = Tab_Cobert
        .Val_IS = Val_IS
        .Per_Franq = Per_Franq
        .Val_Franq_Min = Val_Franq_Min
        .Val_Franq_Max = Val_Franq_Max
        .NUM_PI = NUM_PI
    End With
    Set gCopiaCobert = lobjstcCobert

SAIDA:
    GoSub FechaObj
    Exit Property

FechaObj:
    Set lobjstcCobert = Nothing
    Return
End Property

Public Property Set ObjPai(pobjstc As stcA46V702B)
    Set ObjStcPai = pobjstc
End Property

Public Property Get ObjPai() As stcA46V702B
    Set ObjPai = ObjStcPai
End Property

Private Property Get ObjTabPed() As stcA46V702B
    Set ObjTabPed = ObjStcPai
End Property

Public Property Get Cod_Unicidade() As Long
    If Not ObjStcPai Is Nothing Then
        Cod_Unicidade = ObjStcPai.Cod_Unicidade
    End If
End Property
Public Property Get gPrazo() As Integer
    Static lsRetorno As Integer
    Static lsFlagCalculado As Boolean

    If lsFlagCalculado = False Then
        lsFlagCalculado = True
        lsRetorno = 0
        If Num_Meses_PI <> 0 Then
            Call mpPrazoTabCoefPI(lsRetorno)
        End If
    End If

Retorno:
    gPrazo = lsRetorno
End Property
Private Sub mpPrazoTabCoefPI(ByRef pPrazo As Integer)
    Dim lBDP00Multi As ADODB.Connection
    Dim lMensagem As String
    Dim lrst As clsYasRecordSet
    Dim lSelect As String

    If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, _
                             lBDP00Multi, lMensagem) = False Then
        Call M46V999.gpGraLog(1, lMensagem)
        GoTo SAIDA
    End If
    lSelect = " Select Prazo from Tbm_COEF_PI WITH (NOLOCK) Where"
    lSelect = lSelect & " Cod_Ramo = " & ObjTabPed.Cod_Ramo
    lSelect = lSelect & " and Cod_Cobert = " & Cod_Cobert
    lSelect = lSelect & " and Num_meses = " & Num_Meses_PI
    lSelect = lSelect & " and Cod_Tarif = " & Tab_Cobert
    lSelect = lSelect & " and Cod_Sequencia <= " & ObjTabPed.Cod_Sequencia

    If M46V999.gfObtRegistro(lBDP00Multi, lSelect, lrst, lMensagem) = False Then
        Call M46V999.gpGraLog(1, lMensagem)
        pPrazo = 0
        GoTo SAIDA
    End If
    pPrazo = 0
    If Not lrst.EOF Then
        pPrazo = IIf(IsNull(lrst("Prazo")), 0, lrst("Prazo"))
    End If

    'Supermercado SMART
    If ObjTabPed.Cod_Ramo = 112 And ObjTabPed.Cod_Produto >= 1042 And (ObjTabPed.Cod_Estip >= 2178 And ObjTabPed.Cod_Estip <= 2182) And Cod_Cobert = M46V999.gfBuscarDePara(39, 1) Then
        pPrazo = 7
    End If


SAIDA:
    GoSub FechaObj
    Exit Sub
FechaObj:
    Call M46V999.gpFecha2(lBDP00Multi)
    Set lrst = Nothing
    Set lBDP00Multi = Nothing
    Return
End Sub


Public Function getRpt() As stcA46V732rpt
    Static lrpt As stcA46V732rpt

    If lrpt Is Nothing Then
        Set lrpt = New stcA46V732rpt
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

Public Property Get M46V111() As clsM46V111
    Set M46V111 = M46V999.M46V111
End Property

Private Sub Class_Initialize()
    Set mClsM46V999 = Nothing
End Sub

Private Sub Class_Terminate()
    Set mClsM46V999 = Nothing
End Sub





