VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcPlanoItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mStcPedItem As stcA46V708B


Private Sub Class_Initialize()
    Set mStcPedItem = Nothing
End Sub

Private Sub Class_Terminate()
    Set mStcPedItem = Nothing
End Sub

Public Property Set stcOrigem(ByRef pStcPedItem As stcA46V708B)
    Set mStcPedItem = pStcPedItem
End Property


Public Property Get ValorRegra(ByVal pCodRegra As eRegraPlano) As Double
    Dim lretorno As Double

    On Error GoTo TrataErro: lretorno = 0

    If mStcPedItem.ObjPai.getPlano.TemValorRegra(pCodRegra) = True Then
        lretorno = mStcPedItem.ObjPai.getPlano.ValorRegra(pCodRegra)
    End If
    'Regra de Desconto Comercial
    If pCodRegra = eRegraPlano_DescontoComercial Then
        '1000744 - Item 4 - Plano Mercedes
        If (mStcPedItem.Cod_Plano >= 2021 And mStcPedItem.Cod_Plano <= 2031) And (mStcPedItem.ObjPai.Cod_Produto < 992) Then
            If Trim(mStcPedItem.Cod_Clasf) = mStcPedItem.M46V999.gfBuscarDePara("001", 2) Then
                Select Case Trim(mStcPedItem.Cod_Compl_Clasf)
                Case mStcPedItem.M46V999.gfBuscarDePara("001", 2, "11"): lretorno = 33
                Case mStcPedItem.M46V999.gfBuscarDePara("001", 2, "21"): lretorno = 20
                End Select
            End If
        End If

        '1200903 - Tarifa - 1046 e 1047 - Item 13
        If mStcPedItem.Cod_Plano >= 2051 And mStcPedItem.Cod_Plano <= 2060 And mStcPedItem.Cod_Iden_Tarif < 1046 Then
            If Trim(mStcPedItem.Cod_Clasf) = mStcPedItem.M46V999.gfBuscarDePara("001", 2) Then
                Select Case Trim(mStcPedItem.Cod_Compl_Clasf)
                Case mStcPedItem.M46V999.gfBuscarDePara("001", 2, "11"): lretorno = 20
                Case mStcPedItem.M46V999.gfBuscarDePara("001", 2, "21"): lretorno = 5
                End Select
            End If
        End If

        'Fichas 901282 e 901346
        If mStcPedItem.Cod_Plano >= 2061 And mStcPedItem.Cod_Plano <= 2100 Then
            If Trim(mStcPedItem.Cod_Clasf) = mStcPedItem.M46V999.gfBuscarDePara("001", 2) Then
                Select Case Trim(mStcPedItem.Cod_Compl_Clasf)
                Case mStcPedItem.M46V999.gfBuscarDePara("001", 2, "11"): lretorno = 20
                Case mStcPedItem.M46V999.gfBuscarDePara("001", 2, "21"): lretorno = 5
                End Select
            End If
        End If

        'Fichas 901673 e 901772
        If mStcPedItem.Cod_Plano >= 2111 And mStcPedItem.Cod_Plano <= 2130 Then
            If Trim(mStcPedItem.Cod_Clasf) = mStcPedItem.M46V999.gfBuscarDePara("001", 2) Then
                Select Case Trim(mStcPedItem.Cod_Compl_Clasf)
                Case mStcPedItem.M46V999.gfBuscarDePara("001", 2, "11"): lretorno = 20
                Case mStcPedItem.M46V999.gfBuscarDePara("001", 2, "21"): lretorno = 5
                End Select
            End If
        End If
    End If

SAIDA:
    ValorRegra = lretorno
    Exit Property
TrataErro:
    lretorno = 0
End Property




