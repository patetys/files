VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V715Brpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mStcA46V715B As stcA46V715B

Public Property Set stcOrigem(ByRef pStcA46V715B As stcA46V715B)
    Set mStcA46V715B = pStcA46V715B
End Property

Public Property Get TipBem() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = mStcA46V715B.gTipBem
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    TipBem = lretorno
End Property

Public Property Get DescBen() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = mStcA46V715B.Dsc_Bem
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    DescBen = lretorno
End Property

Public Property Get Val_Bem() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    If mStcA46V715B.Val_Bem > 0 Then
    lretorno = "R$ " & Format(mStcA46V715B.Val_Bem, "standard")
    End If
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    Val_Bem = lretorno
End Property

'Public Property Get XXXXX() As String
'Dim lRetorno As String
''    Select Case tipoRelatorio
''        Case eTipCalcRPT_CotacIndividual
'            lRetorno = mStcA46V715B.WWWWW
''        Case eTipCalcRPT_CotacColetiva
''        Case eTipCalcRPT_PropIndividual
''        Case eTipCalcRPT_PropColetiva
''    End Select
'    lRetorno = IIf(lRetorno = "", " ", lRetorno)
'    XXXXX = lRetorno
'End Property

Private Sub Class_Initialize()
    Set mStcA46V715B = Nothing
End Sub

Private Sub Class_Terminate()
    Set mStcA46V715B = Nothing
End Sub
