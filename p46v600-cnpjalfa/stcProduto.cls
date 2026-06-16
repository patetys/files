VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mStcPed As stcA46V702B
Private mProTipProtocolo As Byte         'Produto: tipo do protocolo.
Private mProNome As String    'Produto: nome.
Private mProVigInicio As Double    'Produto: vigência - início.
Private mProVigTermino As Double  'Produto: vigência - término.
Private mProLimComissao As Double      'Produto: limite da comissão.
Private mProComBasica As Double    'Produto: comissão básica.
Private mProCusApolice As Double     'Produto: custo de apólice.
Private mProPreMinParcela As Double        'Produto: prêmio mínimo por parcela. (no Auto é prêmio mínimo do item)
Private mProPreMinParcelaEndosso As Double           'Produto: prêmio mínimo por parcela do endosso.
Private mProPreMinParcelaTAB As Double
Private mProPremioMinimoTAB As Double
Private mProPremioMinimo As Double       'Produto: prêmio mínimo do item. (no Auto é prêmio mínimo da parcela)
Private mProPreMinEndosso As Double        'Produto: prêmio mínimo por parcela de endosso.
Private mProLimMarComercial As Single          'Produto: limite da margem comercial.
Private mProQuaMaxParcelas As Byte           'Produto: quantidade máxima de parcelas.
Private mProQuaMaxParcelasTAB As Byte
'Private mProJurMinAnt01      As Single           'Produto: limite de juros.
'Private mProJurMinAnt02      As Single           'Produto: limite de juros.
'Private mProJurMinAnt03      As Single           'Produto: limite de juros.
'Private mProJurMinAnt04      As Single           'Produto: limite de juros.
'Private mProJurMinAnt05      As Single           'Produto: limite de juros.
'Private mProJurMinAnt06      As Single           'Produto: limite de juros.
'Private mProJurMinAnt07      As Single           'Produto: limite de juros.
'Private mProJurMinAnt08      As Single           'Produto: limite de juros.
'Private mProJurMinAnt09      As Single           'Produto: limite de juros.
'Private mProJurMinAnt10      As Single           'Produto: limite de juros.
'Private mProJurMinAnt11      As Single           'Produto: limite de juros.
'Private mProJurMinAnt12      As Single           'Produto: limite de juros.
'Private mProJurMinPos01      As Single           'Produto: limite de juros.
'Private mProJurMinPos02      As Single           'Produto: limite de juros.
'Private mProJurMinPos03      As Single           'Produto: limite de juros.
'Private mProJurMinPos04      As Single           'Produto: limite de juros.
'Private mProJurMinPos05      As Single           'Produto: limite de juros.
'Private mProJurMinPos06      As Single           'Produto: limite de juros.
'Private mProJurMinPos07      As Single           'Produto: limite de juros.
'Private mProJurMinPos08      As Single           'Produto: limite de juros.
'Private mProJurMinPos09      As Single           'Produto: limite de juros.
'Private mProJurMinPos10      As Single           'Produto: limite de juros.
'Private mProJurMinPos11      As Single           'Produto: limite de juros.
'Private mProJurMinPos12      As Single           'Produto: limite de juros.

Private mProJurMinAnt01() As Single      'Produto: limite de juros.
Private mProJurMinAnt02() As Single      'Produto: limite de juros.
Private mProJurMinAnt03() As Single      'Produto: limite de juros.
Private mProJurMinAnt04() As Single      'Produto: limite de juros.
Private mProJurMinAnt05() As Single      'Produto: limite de juros.
Private mProJurMinAnt06() As Single      'Produto: limite de juros.
Private mProJurMinAnt07() As Single      'Produto: limite de juros.
Private mProJurMinAnt08() As Single      'Produto: limite de juros.
Private mProJurMinAnt09() As Single      'Produto: limite de juros.
Private mProJurMinAnt10() As Single      'Produto: limite de juros.
Private mProJurMinAnt11() As Single      'Produto: limite de juros.
Private mProJurMinAnt12() As Single      'Produto: limite de juros.
Private mProJurMinPos01() As Single      'Produto: limite de juros.
Private mProJurMinPos02() As Single      'Produto: limite de juros.
Private mProJurMinPos03() As Single      'Produto: limite de juros.
Private mProJurMinPos04() As Single      'Produto: limite de juros.
Private mProJurMinPos05() As Single      'Produto: limite de juros.
Private mProJurMinPos06() As Single      'Produto: limite de juros.
Private mProJurMinPos07() As Single      'Produto: limite de juros.
Private mProJurMinPos08() As Single      'Produto: limite de juros.
Private mProJurMinPos09() As Single      'Produto: limite de juros.
Private mProJurMinPos10() As Single      'Produto: limite de juros.
Private mProJurMinPos11() As Single      'Produto: limite de juros.
Private mProJurMinPos12() As Single      'Produto: limite de juros.


Private mProJurMinPadrao As Single
Private mProJurIOF As Double     'Produto: juros de IOF.
Private mCodProduto As Integer    'Código do TARIFA
Private mCodProtocolo As Long      'Código do protocolo.
Private mCodSequencia As Long  'Código da sequência.
Private mProCusEndosso As Double     'Produto: custo de endosso.

' Thor Hammer - Para evitar carregar a mesma coisa milhares de vezes, a variavel abaixo controla
' o carregamento para que seja feito apenas uma vez. - Daniel Ribeiro 03/11/2015
Private mThor As String
' /Thor Hammer

Private mClsM46V999 As clsM46V999

'===================================================================================================================================
'===================================================================================================================================

Public Property Get Class_Name() As String
    Class_Name = "stcProduto"
End Property

Private Sub Class_Initialize()
'1. Inicializa variáveis - dados do produto.
    Dim lI As Byte

    mProTipProtocolo = 1
    mProNome = " "
    mProVigInicio = 0
    mProVigTermino = 0
    mProLimComissao = 0
    mProComBasica = 0
    mProCusApolice = 0
    mProPreMinParcela = 0
    mProPreMinParcelaTAB = 0
    mProPremioMinimoTAB = 0
    mProPremioMinimo = 0
    mProPreMinEndosso = 0
    mProLimMarComercial = 0
    mProQuaMaxParcelas = 0
    mProQuaMaxParcelasTAB = 0



    mProJurMinPadrao = 3
    mProJurIOF = 0
    mCodProduto = 0
    mCodProtocolo = 0
    mCodSequencia = 0
    mProCusEndosso = 0

    Set mStcPed = Nothing
    Set mClsM46V999 = Nothing

End Sub

Private Sub Class_Terminate()
    Set mClsM46V999 = Nothing
    Set mStcPed = Nothing
End Sub

Public Property Set M46V999(pClsM46V999 As clsM46V999)
    Set mClsM46V999 = pClsM46V999
    Dim lI As Integer

    ReDim mProJurMinAnt01(M46V999.gQtdMaxComp)
    ReDim mProJurMinAnt02(M46V999.gQtdMaxComp)
    ReDim mProJurMinAnt03(M46V999.gQtdMaxComp)
    ReDim mProJurMinAnt04(M46V999.gQtdMaxComp)
    ReDim mProJurMinAnt05(M46V999.gQtdMaxComp)
    ReDim mProJurMinAnt06(M46V999.gQtdMaxComp)
    ReDim mProJurMinAnt07(M46V999.gQtdMaxComp)
    ReDim mProJurMinAnt08(M46V999.gQtdMaxComp)
    ReDim mProJurMinAnt09(M46V999.gQtdMaxComp)
    ReDim mProJurMinAnt10(M46V999.gQtdMaxComp)
    ReDim mProJurMinAnt11(M46V999.gQtdMaxComp)
    ReDim mProJurMinAnt12(M46V999.gQtdMaxComp)
    ReDim mProJurMinPos01(M46V999.gQtdMaxComp)
    ReDim mProJurMinPos02(M46V999.gQtdMaxComp)
    ReDim mProJurMinPos03(M46V999.gQtdMaxComp)
    ReDim mProJurMinPos04(M46V999.gQtdMaxComp)
    ReDim mProJurMinPos05(M46V999.gQtdMaxComp)
    ReDim mProJurMinPos06(M46V999.gQtdMaxComp)
    ReDim mProJurMinPos07(M46V999.gQtdMaxComp)
    ReDim mProJurMinPos08(M46V999.gQtdMaxComp)
    ReDim mProJurMinPos09(M46V999.gQtdMaxComp)
    ReDim mProJurMinPos10(M46V999.gQtdMaxComp)
    ReDim mProJurMinPos11(M46V999.gQtdMaxComp)
    ReDim mProJurMinPos12(M46V999.gQtdMaxComp)

    For lI = 1 To M46V999.gQtdMaxComp
        If lI = 4 Then
            mProJurMinAnt01(lI) = 0
            mProJurMinAnt02(lI) = 0
            mProJurMinAnt03(lI) = 0
            mProJurMinAnt04(lI) = 0
            mProJurMinAnt05(lI) = 0
            mProJurMinAnt06(lI) = 0
            mProJurMinAnt07(lI) = 0
            mProJurMinAnt08(lI) = 0
            mProJurMinAnt09(lI) = 0
            mProJurMinAnt10(lI) = 0
            mProJurMinAnt11(lI) = 0
            mProJurMinAnt12(lI) = 0
            mProJurMinPos01(lI) = 0
            mProJurMinPos02(lI) = 0
            mProJurMinPos03(lI) = 0
            mProJurMinPos04(lI) = 0
            mProJurMinPos05(lI) = 0
            mProJurMinPos06(lI) = 0
            mProJurMinPos07(lI) = 0
            mProJurMinPos08(lI) = 0
            mProJurMinPos09(lI) = 0
            mProJurMinPos10(lI) = 0
            mProJurMinPos11(lI) = 0
            mProJurMinPos12(lI) = 0
        Else
            mProJurMinAnt01(lI) = 3
            mProJurMinAnt02(lI) = 3
            mProJurMinAnt03(lI) = 3
            mProJurMinAnt04(lI) = 3
            mProJurMinAnt05(lI) = 3
            mProJurMinAnt06(lI) = 3
            mProJurMinAnt07(lI) = 3
            mProJurMinAnt08(lI) = 3
            mProJurMinAnt09(lI) = 3
            mProJurMinAnt10(lI) = 3
            mProJurMinAnt11(lI) = 3
            mProJurMinAnt12(lI) = 3
            mProJurMinPos01(lI) = 3
            mProJurMinPos02(lI) = 3
            mProJurMinPos03(lI) = 3
            mProJurMinPos04(lI) = 3
            mProJurMinPos05(lI) = 3
            mProJurMinPos06(lI) = 3
            mProJurMinPos07(lI) = 3
            mProJurMinPos08(lI) = 3
            mProJurMinPos09(lI) = 3
            mProJurMinPos10(lI) = 3
            mProJurMinPos11(lI) = 3
            mProJurMinPos12(lI) = 3
        End If
    Next


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

Public Property Set stcOrigem(ByRef pStcPed As stcA46V702B)
    Set mStcPed = pStcPed
End Property

'===================================================================================================================================
'===================================================================================================================================

Public Property Get ProTipProtocolo() As Byte
    Call gpAtualizaProduto
    ProTipProtocolo = mProTipProtocolo
End Property

Public Property Get ProNome() As String
    Call gpAtualizaProduto
    ProNome = mProNome
End Property

Public Property Get ProVigInicio() As Double
    Call gpAtualizaProduto
    ProVigInicio = mProVigInicio
End Property

Public Property Get ProVigTermino() As Double
    Call gpAtualizaProduto
    ProVigTermino = mProVigTermino
End Property

Public Property Get ProLimComissao() As Double
    Call gpAtualizaProduto
    ProLimComissao = mProLimComissao
End Property

Public Property Get ProComBasica() As Double
    Call gpAtualizaProduto
    ProComBasica = mProComBasica
End Property

Public Property Get ProCusApolice() As Double
    Call gpAtualizaProduto

    If mStcPed.getPlano.EhPlaninho = True Then
        If mStcPed.getPlano.TemValorRegra(eRegraPlano_CustoApolice) = True Then
            mProCusApolice = mStcPed.getPlano.ValorRegra(eRegraPlano_CustoApolice)
        End If
    Else
        '1100237 - Tarifa 1008 e 1009 Abril 2011
        If (mStcPed.Cod_Produto >= 1008) And (mStcPed.Cod_Ramo = e_CodRamo_Residencial) Then
            If mStcPed.Cod_Produto >= 1034 Then
                Select Case mStcPed.gValPrLiqTotal
                Case Is <= 50: mProCusApolice = 30
                Case Is <= 250: mProCusApolice = 50
                Case Is <= 500: mProCusApolice = 60
                Case Is > 500: mProCusApolice = 60
                End Select
            Else
                Select Case mStcPed.gValPrLiqTotal
                Case Is <= 50: mProCusApolice = 30
                Case Is <= 250: mProCusApolice = 50
                Case Is <= 500: mProCusApolice = 75
                Case Is > 500: mProCusApolice = 100
                End Select
            End If
        End If
    End If

    ProCusApolice = mProCusApolice

End Property

Public Property Get ProPreMinParcela() As Double
    Call gpAtualizaProduto

    If mStcPed.getPlano.EhPlaninho = True Then
        If mStcPed.getPlano.TemValorRegra(eRegraPlano_ParcelaMinima) = True Then
            mProPreMinParcela = mStcPed.getPlano.ValorRegra(eRegraPlano_ParcelaMinima)
        End If
    End If
    ProPreMinParcela = mProPreMinParcela
End Property

Public Property Get ProPreMinParcelaEndosso() As Double
    Call gpAtualizaProduto

    ProPreMinParcelaEndosso = mProPreMinParcelaEndosso
End Property

Public Property Get ProPreMinParcelaTAB() As Double
    Call gpAtualizaProduto
    ProPreMinParcelaTAB = mProPreMinParcelaTAB
End Property

Public Property Get ProPremioMinimoTAB() As Double
    Call gpAtualizaProduto
    ProPremioMinimoTAB = mProPremioMinimoTAB
End Property

Public Property Get ProPremioMinimo() As Double
    Dim lQtdiasFalta As Integer
    
    lQtdiasFalta = M46V999.M46V103.mfDifDatas(mStcPed.Dat_Ini_Vig, mStcPed.Dat_Ter_Vig)
    
    If mStcPed.getPlano.EhPlaninho = True Then
        If mStcPed.getPlano.TemValorRegra(eRegraPlano_PremioMinimoPlano) = True Then
            mProPremioMinimo = mStcPed.getPlano.ValorRegra(eRegraPlano_PremioMinimoPlano)
        End If
        
        'If lQtdiasFalta < 365 And mStcPed.Cod_Estip = 20 Then
            'mProPremioMinimo = lQtdiasFalta * Round(mProPremioMinimo / 365, 5)
        'End If
    End If

    Call gpAtualizaProduto
    ProPremioMinimo = mProPremioMinimo
End Property

Public Property Get ProPreMinEndosso() As Double
    Dim lQtdiasFalta As Integer
    
    lQtdiasFalta = M46V999.M46V103.mfDifDatas(mStcPed.Dat_Ini_Vig, mStcPed.Dat_Ter_Vig)
    
    If mStcPed.getPlano.EhPlaninho = True Then
        If mStcPed.getPlano.TemValorRegra(eRegraPlano_PremioMinimoEndosso) = True Then
            mProPreMinEndosso = mStcPed.getPlano.ValorRegra(eRegraPlano_PremioMinimoEndosso)
        End If
        
        'If lQtdiasFalta < 365 And mStcPed.Cod_Estip = 20 Then
            'mProPreMinEndosso = lQtdiasFalta * Round(mProPreMinEndosso / 365, 5)
        'End If
    End If
    
    Call gpAtualizaProduto
    ProPreMinEndosso = mProPreMinEndosso
End Property

Public Property Get ProLimMarComercial() As Single
    Call gpAtualizaProduto
    ProLimMarComercial = mProLimMarComercial
End Property

Public Property Get ProQuaMaxParcelas() As Byte
    Call gpAtualizaProduto

    If mStcPed.getPlano.EhPlaninho = True And mStcPed.Cod_Estip <> 20 Then
        If mStcPed.getPlano.TemValorRegra(eRegraPlano_QtdeParcelas) = True Then
            mProQuaMaxParcelas = mStcPed.getPlano.ValorRegra(eRegraPlano_QtdeParcelas)
        End If
        If InStr(UCase(mStcPed.getPlano.NomPlano), "MERCEDES") Or _
           InStr(UCase(mStcPed.getPlano.NomPlano), "RVA") Then
            mProQuaMaxParcelas = 11
        End If
    ElseIf mStcPed.Cod_Estip = 20 Then
        mProQuaMaxParcelas = 12
    Else
        If M46V999.gOrigemChamada = e_Origem_Kit Then
            M46V999.gParcMaxCorr = 11
        End If
            
        If M46V999.gParcMaxCorr <> 0 And CodProduto >= 1086 Then
            mProQuaMaxParcelas = M46V999.gParcMaxCorr
        End If

    End If
        
    ProQuaMaxParcelas = mProQuaMaxParcelas

End Property

Public Property Get ProQuaMaxParcelasTAB() As Byte
    Call gpAtualizaProduto
    ProQuaMaxParcelasTAB = mProQuaMaxParcelasTAB
End Property

Public Property Get ProJurMinAnt01(ByVal index As Integer) As Single
    Call gpAtualizaProduto
    ProJurMinAnt01 = mProJurMinAnt01(index)
End Property

Public Property Get ProJurMinAnt02(ByVal index As Integer) As Single
    Call gpAtualizaProduto
    ProJurMinAnt02 = mProJurMinAnt02(index)
End Property

Public Property Get ProJurMinAnt03(ByVal index As Integer) As Single
    Call gpAtualizaProduto
    ProJurMinAnt03 = mProJurMinAnt03(index)
End Property

Public Property Get ProJurMinAnt04(ByVal index As Integer) As Single
    Call gpAtualizaProduto
    ProJurMinAnt04 = mProJurMinAnt04(index)
End Property

Public Property Get ProJurMinAnt05(ByVal index As Integer) As Single
    Call gpAtualizaProduto
    ProJurMinAnt05 = mProJurMinAnt05(index)
End Property

Public Property Get ProJurMinAnt06(ByVal index As Integer) As Single
    Call gpAtualizaProduto
    ProJurMinAnt06 = mProJurMinAnt06(index)
End Property

Public Property Get ProJurMinAnt07(ByVal index As Integer) As Single
    Call gpAtualizaProduto
    ProJurMinAnt07 = mProJurMinAnt07(index)
End Property

Public Property Get ProJurMinAnt08(ByVal index As Integer) As Single
    Call gpAtualizaProduto
    ProJurMinAnt08 = mProJurMinAnt08(index)
End Property

Public Property Get ProJurMinAnt09(ByVal index As Integer) As Single
    Call gpAtualizaProduto
    ProJurMinAnt09 = mProJurMinAnt09(index)
End Property

Public Property Get ProJurMinAnt10(ByVal index As Integer) As Single
    Call gpAtualizaProduto
    ProJurMinAnt10 = mProJurMinAnt10(index)
End Property

Public Property Get ProJurMinAnt11(ByVal index As Integer) As Single
    Call gpAtualizaProduto
    ProJurMinAnt11 = mProJurMinAnt11(index)
End Property

Public Property Get ProJurMinAnt12(ByVal index As Integer) As Single
    Call gpAtualizaProduto
    ProJurMinAnt12 = mProJurMinAnt12(index)
End Property

Public Property Get ProJurMinPos01(ByVal index As Integer) As Single
    Call gpAtualizaProduto
    ProJurMinPos01 = mProJurMinPos01(index)
End Property

Public Property Get ProJurMinPos02(ByVal index As Integer) As Single
    Call gpAtualizaProduto
    ProJurMinPos02 = mProJurMinPos02(index)
End Property

Public Property Get ProJurMinPos03(ByVal index As Integer) As Single
    Call gpAtualizaProduto
    ProJurMinPos03 = mProJurMinPos03(index)
End Property

Public Property Get ProJurMinPos04(ByVal index As Integer) As Single
    Call gpAtualizaProduto
    ProJurMinPos04 = mProJurMinPos04(index)
End Property

Public Property Get ProJurMinPos05(ByVal index As Integer) As Single
    Call gpAtualizaProduto
    ProJurMinPos05 = mProJurMinPos05(index)
End Property

Public Property Get ProJurMinPos06(ByVal index As Integer) As Single
    Call gpAtualizaProduto
    ProJurMinPos06 = mProJurMinPos06(index)
End Property

Public Property Get ProJurMinPos07(ByVal index As Integer) As Single
    Call gpAtualizaProduto
    ProJurMinPos07 = mProJurMinPos07(index)
End Property

Public Property Get ProJurMinPos08(ByVal index As Integer) As Single
    Call gpAtualizaProduto
    ProJurMinPos08 = mProJurMinPos08(index)
End Property

Public Property Get ProJurMinPos09(ByVal index As Integer) As Single
    Call gpAtualizaProduto
    ProJurMinPos09 = mProJurMinPos09(index)
End Property

Public Property Get ProJurMinPos10(ByVal index As Integer) As Single
    Call gpAtualizaProduto
    ProJurMinPos10 = mProJurMinPos10(index)
End Property

Public Property Get ProJurMinPos11(ByVal index As Integer) As Single
    Call gpAtualizaProduto
    ProJurMinPos11 = mProJurMinPos11(index)
End Property

Public Property Get ProJurMinPos12(ByVal index As Integer) As Single
    Call gpAtualizaProduto
    ProJurMinPos12 = mProJurMinPos12(index)
End Property

Public Property Get ProJurMinPadrao() As Single
'tbm_juros
    Call gpAtualizaProduto
    ProJurMinPadrao = mProJurMinPadrao
End Property

Public Property Get ProJurIOF() As Double
    Call gpAtualizaProduto
    ProJurIOF = mProJurIOF
End Property

Public Property Get CodProduto() As Integer
    Call gpAtualizaProduto
    CodProduto = mCodProduto
End Property

Public Property Get CodProtocolo() As Long
    Call gpAtualizaProduto
    CodProtocolo = mCodProtocolo
End Property

Public Property Get CodSequencia() As Long
    Call gpAtualizaProduto
    CodSequencia = mCodSequencia
End Property

Public Property Get ProCusEndosso() As Double
    Call gpAtualizaProduto
    ProCusEndosso = mProCusEndosso
End Property

'===================================================================================================================================
'===================================================================================================================================
Private Sub gpAtualizaProduto()
    Static lsCodProduto As Integer
    Static lsCodSequencia As Long
    Static lsCodRamo As Long
    Static lsModalidade As Double               'Código do estipulante.
    Dim lMsgErr As String

    If (lsCodProduto <> mStcPed.Cod_Produto) Or _
       (lsModalidade <> mStcPed.Cod_Estip) Or _
       (lsCodSequencia <> CLng(mStcPed.Cod_Sequencia)) Or _
       (lsCodRamo <> CLng(mStcPed.Cod_Ramo)) Then

        lsCodProduto = mStcPed.Cod_Produto
        lsCodSequencia = CLng(mStcPed.Cod_Sequencia)
        lsCodRamo = CLng(mStcPed.Cod_Ramo)
        lsModalidade = mStcPed.Cod_Estip

        Call mpObtProduto(lsCodProduto, lsCodSequencia, lsCodRamo, lsModalidade, mStcPed, lMsgErr)
    End If

    Call mpAtualizaGlobais

End Sub

Private Sub mpObtProduto(ByVal pCodProduto As Integer, ByVal pCodSequencia As Long, ByVal pCodRamo As Long, ByVal pCodModalidade As Integer, ByRef pStcPed As stcA46V702B, ByRef pMsgErr As String)
'Procedure: obtém produto e informações dependentes de produto.

'Parâmetros de entrada...pCodProduto.....código do produto.
'                        pCodProtocolo...código do protocolo.
'                        pCodSequencia...código da seqüência.

    Dim lbdTAB_PROD As ADODB.Connection    'Base de dados - TAB_PROD.
    Dim lMensagem As String     'Mensagem.
    Dim lrsTAB_PROD As clsYasRecordSet    'Registro - TAB_PROD.
    Dim lSelect As String
    Dim lTipPag As Integer

    '2. Se produto zerado, encerra processo.
    If pCodProduto = 0 Or pCodRamo = 0 Then
        Exit Sub
    End If
    
    '2.a. Se o produto já foi carregado, não precisa carregar de novo - Thor Hammer
    If pCodRamo & "." & pCodModalidade & "." & pCodProduto = mThor Then
        Exit Sub
    End If

    '3. Abre arquivo.
    If mStcPed.M46V999.gfAbrBasDados("P00MULTI", mStcPed.M46V999.gTipP00MULTI, mStcPed.M46V999.gEndP00MULTI, lbdTAB_PROD, lMensagem) = False Then
        pMsgErr = lMensagem
        Exit Sub
    End If

    '4. Obtém produto informado.
    lSelect = "SELECT * FROM TAB_PRODUTO WITH (NOLOCK) WHERE Cod_Produto = " & pCodProduto & " AND Cod_Ramo = " & _
              pCodRamo & " and cod_sequencia = " & pCodSequencia
    If mStcPed.M46V999.gfObtRegistro(lbdTAB_PROD, lSelect, lrsTAB_PROD, lMensagem) = False Then
        Call mStcPed.M46V999.gpFecha2(lbdTAB_PROD)
        pMsgErr = lMensagem
        Exit Sub
    End If
    If lrsTAB_PROD.EOF = True Then
        Call mStcPed.M46V999.gpFecha2(lbdTAB_PROD)
        pMsgErr = "FNL0002 - " & mStcPed.M46V999.gPrefixo & "Produto não localizado (" & _
                  Format$(pCodProduto, "000") & "/" & Format$(pCodRamo, "000") & ")."
        Exit Sub
    End If

    '5. Carrega variáveis globais.
    mCodProduto = lrsTAB_PROD!Cod_Produto
    mCodProtocolo = lrsTAB_PROD!Cod_Protocolo
    mCodSequencia = lrsTAB_PROD!Cod_Sequencia
    mProTipProtocolo = lrsTAB_PROD!Tip_Prot
    mProNome = lrsTAB_PROD!Nom_Prot
    mProVigInicio = lrsTAB_PROD!Dat_Vig_Ini
    mProVigTermino = lrsTAB_PROD!Dat_Vig_Ter
    mProLimComissao = lrsTAB_PROD!Per_Cms_Limite
    mProComBasica = lrsTAB_PROD!Per_Cms_Basica
    mProCusEndosso = lrsTAB_PROD!Val_Cus_Emis_Endo
    mProCusApolice = lrsTAB_PROD!Val_Cus_Emis_Apol
    mProPreMinParcela = lrsTAB_PROD!Par_Min_Parcela
    mProPreMinParcelaEndosso = lrsTAB_PROD!Par_Min_Endosso
    mProPreMinParcelaTAB = mProPreMinParcela
    mProPremioMinimo = lrsTAB_PROD!Pre_Min_Parcela
    mProPremioMinimoTAB = mProPremioMinimo
    mProPreMinEndosso = lrsTAB_PROD!Pre_Min_Endosso
    mProLimMarComercial = lrsTAB_PROD!Lim_Marg_Coml
    mProQuaMaxParcelas = lrsTAB_PROD!Qua_Max_Parcelas
    mProQuaMaxParcelasTAB = lrsTAB_PROD!Qua_Max_Parcelas


    mProJurMinPadrao = lrsTAB_PROD!Jur_Padrao
    mProJurIOF = lrsTAB_PROD!Jur_IOF

    Call mStcPed.M46V999.gpFecha2(lbdTAB_PROD)

    'Substituir por chamado do TBM_JUROS
    '6. Obtém os juros para esse ramo e produto

    lSelect = "SELECT * FROM TBM_JUROS WITH (NOLOCK)  WHERE "
    lSelect = lSelect & "  Cod_Ramo = " & pCodRamo
    lSelect = lSelect & " and Cod_Modalidade in (0," & pCodModalidade & ")"
    lSelect = lSelect & " and Cod_Tarifa  = " & pCodProduto
    
    If pStcPed.ColecaoPedCorr.Count > 0 Then
        lSelect = lSelect & " and Cod_Corretor  IN (0, " & pStcPed.ColecaoPedCorr(1).Cod_Corr & ")"
    Else
        lSelect = lSelect & " and Cod_Corretor  = 0"
    End If
    
    lSelect = lSelect & " order by Cod_Corretor, Cod_Modalidade,  Tip_Compensacao,   Num_Parcela "



    If mStcPed.M46V999.gfObtRegistro(lbdTAB_PROD, lSelect, lrsTAB_PROD, lMensagem) = False Then
        Call mStcPed.M46V999.gpFecha2(lbdTAB_PROD)
        pMsgErr = lMensagem
        Exit Sub
    End If
    If lrsTAB_PROD.EOF = True Then
        Call mStcPed.M46V999.gpFecha2(lbdTAB_PROD)
        pMsgErr = "FNL0002 - " & mStcPed.M46V999.gPrefixo & "Produto não localizado (" & _
                  Format$(pCodProduto, "000") & "/" & Format$(pCodRamo, "000") & ")."
        Exit Sub
    End If

    Do While Not lrsTAB_PROD.EOF
        lTipPag = lrsTAB_PROD!Tip_Compensacao
        Select Case lrsTAB_PROD!Num_Parcela
        Case 1
            mProJurMinAnt01(lTipPag) = lrsTAB_PROD!Per_Juros_Antec
            mProJurMinPos01(lTipPag) = lrsTAB_PROD!Per_Juros_Pos
        Case 2
            mProJurMinAnt02(lTipPag) = lrsTAB_PROD!Per_Juros_Antec
            mProJurMinPos02(lTipPag) = lrsTAB_PROD!Per_Juros_Pos
        Case 3
            mProJurMinAnt03(lTipPag) = lrsTAB_PROD!Per_Juros_Antec
            mProJurMinPos03(lTipPag) = lrsTAB_PROD!Per_Juros_Pos
        Case 4
            mProJurMinAnt04(lTipPag) = lrsTAB_PROD!Per_Juros_Antec
            mProJurMinPos04(lTipPag) = lrsTAB_PROD!Per_Juros_Pos
        Case 5
            mProJurMinAnt05(lTipPag) = lrsTAB_PROD!Per_Juros_Antec
            mProJurMinPos05(lTipPag) = lrsTAB_PROD!Per_Juros_Pos
        Case 6
            mProJurMinAnt06(lTipPag) = lrsTAB_PROD!Per_Juros_Antec
            mProJurMinPos06(lTipPag) = lrsTAB_PROD!Per_Juros_Pos
        Case 7
            mProJurMinAnt07(lTipPag) = lrsTAB_PROD!Per_Juros_Antec
            mProJurMinPos07(lTipPag) = lrsTAB_PROD!Per_Juros_Pos
        Case 8
            mProJurMinAnt08(lTipPag) = lrsTAB_PROD!Per_Juros_Antec
            mProJurMinPos08(lTipPag) = lrsTAB_PROD!Per_Juros_Pos
        Case 9
            mProJurMinAnt09(lTipPag) = lrsTAB_PROD!Per_Juros_Antec
            mProJurMinPos09(lTipPag) = lrsTAB_PROD!Per_Juros_Pos
        Case 10
            mProJurMinAnt10(lTipPag) = lrsTAB_PROD!Per_Juros_Antec
            mProJurMinPos10(lTipPag) = lrsTAB_PROD!Per_Juros_Pos
        Case 11
            mProJurMinAnt11(lTipPag) = lrsTAB_PROD!Per_Juros_Antec
            mProJurMinPos11(lTipPag) = lrsTAB_PROD!Per_Juros_Pos
        Case 12
            mProJurMinAnt12(lTipPag) = lrsTAB_PROD!Per_Juros_Antec
            mProJurMinPos12(lTipPag) = lrsTAB_PROD!Per_Juros_Pos
        End Select
        lrsTAB_PROD.MoveNext
    Loop
    
    Call mStcPed.M46V999.gpFecha2(lbdTAB_PROD)
    mThor = pCodRamo & "." & pCodModalidade & "." & pCodProduto
End Sub

Private Sub mpAtualizaGlobais()
    Dim lI As Integer
    Dim lMsg As String

    If mStcPed.M46V999.gPrograma <> "D046O003" And mStcPed.Cod_Estip <> 20 Then Exit Sub

    mStcPed.M46O001.gProTipProtocolo = mProTipProtocolo
    mStcPed.M46O001.gProNome = mProNome
    mStcPed.M46O001.gProVigInicio = mProVigInicio
    mStcPed.M46O001.gProVigTermino = mProVigTermino
    mStcPed.M46O001.gProLimComissao = mProLimComissao
    mStcPed.M46O001.gProComBasica = mProComBasica
    mStcPed.M46V999.gProCusApolice = mProCusApolice
    mStcPed.M46V999.gProPreMinParcela = mProPreMinParcela
    mStcPed.M46V999.gProPreMinParcelaTAB = mProPreMinParcelaTAB
    mStcPed.M46V999.gProPremioMinimoTAB = mProPremioMinimoTAB
    mStcPed.M46V999.gProPremioMinimo = mProPremioMinimo
    mStcPed.M46V999.gProPreMinEndosso = mProPreMinEndosso
    mStcPed.M46O001.gProLimMarComercial = mProLimMarComercial
    If mStcPed.Cod_Estip = 20 Then
        mStcPed.M46V999.gProQuaMaxParcelas = 12
    Else
        mStcPed.M46V999.gProQuaMaxParcelas = mProQuaMaxParcelas
    End If
    mStcPed.M46V999.gProQuaMaxParcelasTAB = mProQuaMaxParcelasTAB

    'tmb_juros
    '    For lI = 1 To mStcPed.M46V999.gQtdMaxComp
    '        mStcPed.M46O001.gProJurMinAnt01(lI) = mProJurMinAnt01(lI)
    '        mStcPed.M46O001.gProJurMinAnt02(lI) = mProJurMinAnt02(lI)
    '        mStcPed.M46O001.gProJurMinAnt03(lI) = mProJurMinAnt03(lI)
    '        mStcPed.M46O001.gProJurMinAnt04(lI) = mProJurMinAnt04(lI)
    '        mStcPed.M46O001.gProJurMinAnt05(lI) = mProJurMinAnt05(lI)
    '        mStcPed.M46O001.gProJurMinAnt06(lI) = mProJurMinAnt06(lI)
    '        mStcPed.M46O001.gProJurMinAnt07(lI) = mProJurMinAnt07(lI)
    '        mStcPed.M46O001.gProJurMinAnt08(lI) = mProJurMinAnt08(lI)
    '        mStcPed.M46O001.gProJurMinAnt09(lI) = mProJurMinAnt09(lI)
    '        mStcPed.M46O001.gProJurMinAnt10(lI) = mProJurMinAnt10(lI)
    '        mStcPed.M46O001.gProJurMinAnt11(lI) = mProJurMinAnt11(lI)
    '        mStcPed.M46O001.gProJurMinAnt12(lI) = mProJurMinAnt12(lI)
    '        mStcPed.M46O001.gProJurMinPos01(lI) = mProJurMinPos01(lI)
    '        mStcPed.M46O001.gProJurMinPos02(lI) = mProJurMinPos02(lI)
    '        mStcPed.M46O001.gProJurMinPos03(lI) = mProJurMinPos03(lI)
    '        mStcPed.M46O001.gProJurMinPos04(lI) = mProJurMinPos04(lI)
    '        mStcPed.M46O001.gProJurMinPos05(lI) = mProJurMinPos05(lI)
    '        mStcPed.M46O001.gProJurMinPos06(lI) = mProJurMinPos06(lI)
    '        mStcPed.M46O001.gProJurMinPos07(lI) = mProJurMinPos07(lI)
    '        mStcPed.M46O001.gProJurMinPos08(lI) = mProJurMinPos08(lI)
    '        mStcPed.M46O001.gProJurMinPos09(lI) = mProJurMinPos09(lI)
    '        mStcPed.M46O001.gProJurMinPos10(lI) = mProJurMinPos10(lI)
    '        mStcPed.M46O001.gProJurMinPos11(lI) = mProJurMinPos11(lI)
    '        mStcPed.M46O001.gProJurMinPos12(lI) = mProJurMinPos12(lI)
    '    Next

    mStcPed.M46O001.gProJurMinPadrao = mProJurMinPadrao
    mStcPed.M46V111.gProJurIOF = mProJurIOF
    mStcPed.M46V999.gCodProduto = mCodProduto
    mStcPed.M46V999.gCodProtocolo = mCodProtocolo
    mStcPed.M46V999.gCodSequencia = mCodSequencia
    mStcPed.M46V999.gProCusEndosso = mProCusEndosso

End Sub

Public Property Get ProQuaMaxParcelasSegmentados() As Byte
    If CodProduto >= 1086 And (M46V999.gParcMaxCorr <> 0 Or M46V999.gPrograma <> "D046O003") Then
        ProQuaMaxParcelasSegmentados = 11
    Else
        ProQuaMaxParcelasSegmentados = 7
    End If

End Property

Public Property Set ParcMaxCorr(lParcMaxCorr)
    mProQuaMaxParcelas = lParcMaxCorr
End Property

Public Property Get ProPreMinCobert() As Double
    Dim lQtdiasFalta As Integer
    Dim mProPreMinCobert As Double
    
    lQtdiasFalta = M46V999.M46V103.mfDifDatas(mStcPed.Dat_Ini_Vig, mStcPed.Dat_Ter_Vig)
    mProPreMinCobert = 0
    
    If mStcPed.getPlano.EhPlaninho = True Then
        If mStcPed.getPlano.TemValorRegra(eRegraPlano_PremioMinimoCobertura) = True Then
            mProPreMinCobert = mStcPed.getPlano.ValorRegra(eRegraPlano_PremioMinimoCobertura)
        End If
        
        If lQtdiasFalta < 365 And mStcPed.Cod_Estip = 20 Then
            mProPreMinCobert = lQtdiasFalta * Round(mProPreMinCobert / 365, 5)
        End If
    End If
    
    ProPreMinCobert = mProPreMinCobert
End Property

Public Property Get ProPreMinEndossoPlano() As Double
    Dim lQtdiasFalta As Integer
    Dim mProPreMinEndossoPlano As Double
    
    lQtdiasFalta = M46V999.M46V103.mfDifDatas(mStcPed.Dat_Ini_Vig, mStcPed.Dat_Ter_Vig)
    mProPreMinEndossoPlano = 0
    
    If mStcPed.getPlano.EhPlaninho = True Then
        If mStcPed.getPlano.TemValorRegra(eRegraPlano_PremioMinimoEndosso) = True Then
            mProPreMinEndossoPlano = mStcPed.getPlano.ValorRegra(eRegraPlano_PremioMinimoEndosso)
        End If
        
        'If lQtdiasFalta < 365 And mStcPed.Cod_Estip = 20 Then
         '   mProPreMinEndossoPlano = lQtdiasFalta * Round(mProPreMinEndossoPlano / 365, 5)
        'End If
    End If
    
    ProPreMinEndossoPlano = mProPreMinEndossoPlano
End Property

Public Property Get ProParcMinApolAntec() As Double
    Dim lQtdiasFalta As Integer
    Dim mProParcMinApolAntec As Double
    
    lQtdiasFalta = M46V999.M46V103.mfDifDatas(mStcPed.Dat_Ini_Vig, mStcPed.Dat_Ter_Vig)
    mProParcMinApolAntec = 0
    
    If mStcPed.getPlano.EhPlaninho = True Then
        If mStcPed.getPlano.TemValorRegra(eRegraPlano_ParcelaMinimaApolAntecipada) = True Then
            mProParcMinApolAntec = mStcPed.getPlano.ValorRegra(eRegraPlano_ParcelaMinimaApolAntecipada)
        End If
        
        'If lQtdiasFalta < 365 And mStcPed.Cod_Estip = 20 Then
         '   mProParcMinApolAntec = lQtdiasFalta * Round(mProParcMinApolAntec / 365, 5)
        'End If
    End If
    
    ProParcMinApolAntec = mProParcMinApolAntec
End Property

Public Property Get ProParcMinApolPostec() As Double
    Dim lQtdiasFalta As Integer
    Dim mProParcMinApolPostec As Double
    
    lQtdiasFalta = M46V999.M46V103.mfDifDatas(mStcPed.Dat_Ini_Vig, mStcPed.Dat_Ter_Vig)
    mProParcMinApolPostec = 0
    
    If mStcPed.getPlano.EhPlaninho = True Then
        If mStcPed.getPlano.TemValorRegra(eRegraPlano_ParcelaMinimaApolPostecipada) = True Then
            mProParcMinApolPostec = mStcPed.getPlano.ValorRegra(eRegraPlano_ParcelaMinimaApolPostecipada)
        End If
        
        'If lQtdiasFalta < 365 And mStcPed.Cod_Estip = 20 Then
        '    mProParcMinApolPostec = lQtdiasFalta * Round(mProParcMinApolPostec / 365, 5)
        'End If
    End If
    
    ProParcMinApolPostec = mProParcMinApolPostec
End Property

Public Property Get ProParcMinEndosso() As Double
    Dim lQtdiasFalta As Integer
    Dim mProParcMinEndosso As Double
    
    lQtdiasFalta = M46V999.M46V103.mfDifDatas(mStcPed.Dat_Ini_Vig, mStcPed.Dat_Ter_Vig)
    mProParcMinEndosso = 0
    
    If mStcPed.getPlano.EhPlaninho = True Then
        If mStcPed.getPlano.TemValorRegra(eRegraPlano_ParcelaMinimaEndosso) = True Then
            mProParcMinEndosso = mStcPed.getPlano.ValorRegra(eRegraPlano_ParcelaMinimaEndosso)
        End If
        
        'If lQtdiasFalta < 365 And mStcPed.Cod_Estip = 20 Then
        '    mProParcMinEndosso = lQtdiasFalta * Round(mProParcMinEndosso / 365, 5)
        'End If
    End If
    
    ProParcMinEndosso = mProParcMinEndosso
End Property
