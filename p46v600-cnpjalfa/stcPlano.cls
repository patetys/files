VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcPlano"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mStcPed As stcA46V702B
Private colecaoRegras As Collection

Private mEhSegmentado As Integer ' 0 = Não definido, 1 = Sim, 2 = Não
Private mIndOcultaWeb As Integer  ' 20170524: Ind_Oculta_Web : 0 = Exibe Plano no Cotador Web, 1 = Oculta Plano

Private Sub Class_Initialize()
    Set colecaoRegras = New Collection
    Set mStcPed = Nothing
    mIndOcultaWeb = 0
End Sub

Private Sub Class_Terminate()
    Set mStcPed = Nothing
    Set colecaoRegras = Nothing
End Sub

Public Property Set stcOrigem(ByRef pStcPed As stcA46V702B)
    Set mStcPed = pStcPed
End Property

'===================================================================================================================================
'===================================================================================================================================
Public Property Get NomPlano() As String
    NomPlano = mfNomePlano
End Property

Public Property Get EhPlaninho() As Boolean
    Call mpCarregaPlaninho
    EhPlaninho = CBool(colecaoRegras.Count > 0)
End Property

Public Property Get EhPlanoSegmentado() As Boolean
    Dim NomePlano As String
    NomePlano = mfNomePlano
    EhPlanoSegmentado = CBool(mEhSegmentado = 1)
End Property

Public Property Get ehPlanoEngessado() As Boolean
    Dim lretorno As Boolean

    If EhPlaninho = True Then
        lretorno = TemValorRegra(eRegraPlano_ValorRisco)
    End If

    ehPlanoEngessado = lretorno

End Property

Public Property Get ehPlanoImobiliaria() As Boolean
    Dim lretorno As Boolean

    If EhPlaninho = True Then
        lretorno = CBool(TemValorRegra(eRegraPlano_ValorRisco) = False)
    End If

    ehPlanoImobiliaria = lretorno

End Property

Public Property Get TemValorRegra(ByVal pCodRegra As eRegraPlano) As Boolean
    Dim lstcDummy As stcRegraPlano
    Dim lretorno As Boolean

    Err.Clear
    On Error GoTo TrataErro: lretorno = False

    If EhPlaninho = True Then
        Set lstcDummy = colecaoRegras.Item(CStr(pCodRegra))
        If lstcDummy.CodRegra = pCodRegra Then
            lretorno = True
        End If
    End If

SAIDA:
    Err.Clear
    TemValorRegra = lretorno
    Exit Property
TrataErro:
    lretorno = False
End Property

Public Property Get ValorRegra(ByVal pCodRegra As eRegraPlano) As Double
    Dim lstcDummy As stcRegraPlano
    Dim lretorno As Double

    Err.Clear
    On Error GoTo TrataErro: lretorno = 0

    If TemValorRegra(pCodRegra) = True Then
        Set lstcDummy = colecaoRegras.Item(CStr(pCodRegra))
        If lstcDummy.CodRegra = pCodRegra Then
            lretorno = lstcDummy.ValRegra
        End If
    End If

SAIDA:
    Err.Clear
    ValorRegra = lretorno
    Exit Property
TrataErro:
    lretorno = 0
End Property

'' 20170524: Indica Plano Oculto Web
Public Property Get IndOcultaWeb() As Integer
    IndOcultaWeb = mIndOcultaWeb
End Property
'===================================================================================================================================
'===================================================================================================================================

Private Sub mpCarregaPlaninho()
    Dim lP00Multi_Tplano As ADODB.Connection
    Dim lrstRegraPlano As clsYasRecordSet
    Dim lMensagem As String
    Dim lSelect As String
    Dim lstcRegraPlano As stcRegraPlano

    Static lsCodRamo As e_CodRamo         'Código do Ramo.
    Static lsModalidade As Double               'Código do estipulante.
    Static lsCodSequencia As Integer
    Static lsCodTarif As Integer            'Código do produto.
    Static lretorno As Boolean

    If (lsCodRamo <> mStcPed.Cod_Ramo) Or _
       (lsModalidade <> mStcPed.Cod_Estip) Or _
       (lsCodSequencia <> mStcPed.Cod_Sequencia) Or _
       (lsCodTarif <> mStcPed.Cod_Produto) Then

        lsCodRamo = mStcPed.Cod_Ramo
        lsModalidade = mStcPed.Cod_Estip
        lsCodSequencia = mStcPed.Cod_Sequencia
        lsCodTarif = mStcPed.Cod_Produto

        Set colecaoRegras = New Collection

        If mStcPed.M46V999.gfAbrBasDados("P00MULTI", mStcPed.M46V999.gTipP00MULTI, mStcPed.M46V999.gEndP00MULTI, lP00Multi_Tplano, lMensagem) = False Then Call mStcPed.M46V999.gpGraLog(2, lMensagem): Exit Sub

        lSelect = "SELECT COD_REGRA,VAL_REGRA FROM TMB_REGRA_PLANO WITH (NOLOCK) "
        lSelect = lSelect & " Where COD_RAMO = " & lsCodRamo
        lSelect = lSelect & " AND COD_MODALIDADE = " & lsModalidade
        lSelect = lSelect & " AND COD_SEQUENCIA <=" & lsCodSequencia
        lSelect = lSelect & " AND COD_TARIF=" & lsCodTarif
        lSelect = lSelect & " Order By Cod_Regra"

        If mStcPed.M46V999.gfObtRegistro(lP00Multi_Tplano, lSelect, lrstRegraPlano, lMensagem) = False Then Call mStcPed.M46V999.gpGraLog(2, lMensagem): Exit Sub

        Do While lrstRegraPlano.EOF = False
            Set lstcRegraPlano = New stcRegraPlano

            With lstcRegraPlano
                .CodRegra = mStcPed.M46V999.gfTratarNulo(lrstRegraPlano("COD_REGRA"), tpoNumerico)
                .ValRegra = mStcPed.M46V999.gfTratarNulo(lrstRegraPlano("val_regra"), tpoNumerico)
            End With
            colecaoRegras.Add lstcRegraPlano, lstcRegraPlano.Class_ID

            lrstRegraPlano.MoveNext
        Loop
        Call mStcPed.M46V999.gpFecha2(lP00Multi_Tplano)
    End If

End Sub

Private Function mfNomePlano() As String
    Dim lCN_RamosDiversos As ADODB.Connection
    Dim lrs As clsYasRecordSet
    Dim lMensagem As String
    Dim lSelect As String

    Static lsCodRamo As e_CodRamo         'Código do Ramo.
    Static lsCodPlano As Double
    Static lretorno As String

    If (lsCodRamo <> mStcPed.Cod_Ramo) Or _
       (lsCodPlano <> mStcPed.Cod_Estip) Then

        lsCodRamo = mStcPed.Cod_Ramo
        lsCodPlano = mStcPed.Cod_Estip

        If mStcPed.M46V999.gfAbrBasDados("P00MULTI", mStcPed.M46V999.gTipP00MULTI, mStcPed.M46V999.gEndP00MULTI, lCN_RamosDiversos, lMensagem) = False Then Call mStcPed.M46V999.gpGraLog(2, lMensagem): Exit Function

        'lSelect = "SELECT Nom_Prot, Segmentado FROM TAB_Plano WITH (NOLOCK) "
        lSelect = "SELECT Nom_Prot, Segmentado, Ind_Oculta_Web FROM TAB_Plano WITH (NOLOCK) "
        lSelect = lSelect & " Where COD_RAMO = " & lsCodRamo
        lSelect = lSelect & " AND Cod_Produto = " & lsCodPlano

        If mStcPed.M46V999.gfObtRegistro(lCN_RamosDiversos, lSelect, lrs, lMensagem) = False Then Call mStcPed.M46V999.gpGraLog(2, lMensagem): Exit Function

        Do While lrs.EOF = False
            lretorno = mStcPed.M46V999.gfTratarNulo(lrs("Nom_Prot"), tpoString)
            mEhSegmentado = IIf(lrs("Segmentado"), 1, 2)
            mIndOcultaWeb = mStcPed.M46V999.gfTratarNulo(lrs("Ind_Oculta_Web"), TpoNumInteiro)
            lrs.MoveNext
        Loop
        Call mStcPed.M46V999.gpFecha2(lCN_RamosDiversos)
    End If

    mfNomePlano = lretorno

End Function

Public Function gVerificaPlanoAtivo() As Boolean
    Dim lCN_RamosDiversos As ADODB.Connection
    Dim lrs As clsYasRecordSet
    Dim lMensagem As String
    Dim lSelect As String

    On Error GoTo TrataErro

    gVerificaPlanoAtivo = False
    
    If mStcPed.M46V999.gfAbrBasDados("P00MULTI", mStcPed.M46V999.gTipP00MULTI, mStcPed.M46V999.gEndP00MULTI, lCN_RamosDiversos, lMensagem) = False Then GoTo TrataErro

    lSelect = "EXEC PROC_ObtemPlaninhoAtivo " & mStcPed.M46V111.gCod_plano 'Obtem o plano da variavel global pois neste momento ainda não tenho na classe
    lSelect = lSelect & ", " & mStcPed.Cod_Ramo
    
    If mStcPed.M46V999.gfObtRegistro(lCN_RamosDiversos, lSelect, lrs, lMensagem) = False Then GoTo TrataErro

    If lrs.RecordCount > 0 Then
       gVerificaPlanoAtivo = True
    End If
SAIDA:
    Set lCN_RamosDiversos = Nothing
    Set lrs = Nothing
    
    Exit Function
TrataErro:
    Call mStcPed.M46V999.gpGraLog(2, lMensagem)
    GoTo SAIDA
End Function
