VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V715B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: stcA46V715B - Tab_Ped_Loc_Bem
' Data/Hora: 22/09/2005 - 15:02:33
' Sistema: Multi
' Thais Juren Gomes
' ======================================================================================================================================================
Public Num_Item As Long         'PK         'Número do item.
Public nosso_numero As String   'Nosso Número
Public Cod_Org As Long    'Código do órgão.
Public Cod_Ramo As e_CodRamo    'Código do ramo.
Public Num_Ped As Long    'Número do pedido.
Public Num_Apol As Double    'Número da apólice.
Public Num_Endo As Long    'Número do endosso.
Public Dsc_Bem As String    'Descrição do Bem.
Public Val_Bem As Double    'Valor do bem.
Public Cod_Bem As Long    'Código do bem.
Public Tip_Bem As String
Public Cod_Seq As Long    'Código de Sequencia.
Public Classe_Bem As String
Public Observacao As String
' ======================================================================================================================================================
Private ObjStcPai As stcA46V708B
' ======================================================================================================================================================

Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "A46V715B"
End Property

Public Property Get gTipBem() As String
    Dim lbdTbm_Tip_Bens As ADODB.Connection    'Arquivo: P00Tarifa.
    Dim lMensagem As String     'Mensagem.
    Dim lrsTbm_Tip_Bens As clsYasRecordSet  'Registro : TAB_COBERT.
    Dim lSelect As String

    If Not ObjStcPai Is Nothing Then
        If ObjStcPai.Cod_Iden_Tarif < 1102 Then
            If M46V999.gfPreenchido(Tip_Bem) Then
                gTipBem = Tip_Bem
            Else
                If Cod_Bem <> 0 Then
                    If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, lbdTbm_Tip_Bens, lMensagem) = False Then
                        Call M46V999.gpGraLog(2, lMensagem)
                        GoTo SAIDA
                    End If
        
                    lSelect = " Select top 1 * from Tbm_Tip_Bens WITH (NOLOCK) "
                    lSelect = lSelect & " where Cod_Bem  = " & Cod_Bem
                    lSelect = lSelect & "   and Cod_tarif  = " & M46V999.gCodProduto
                    lSelect = lSelect & "   and Cod_Sequencia <= " & M46V999.gCodSequencia
        
                    '## smrosa
                    If M46V999.gCodProduto >= 947 Then
                        lSelect = lSelect & "    and Cod_Ramo  = " & M46V999.gCod_Ramo
                    End If
        
                    lSelect = lSelect & "   order by Cod_Sequencia desc"
        
                    If M46V999.gfObtRegistro(lbdTbm_Tip_Bens, lSelect, lrsTbm_Tip_Bens, lMensagem) = False Then
                        Call M46V999.gpGraLog(2, lMensagem)
                        Call M46V999.gpFecha2(lbdTbm_Tip_Bens)
                        GoTo SAIDA
                    End If
                    If Not lrsTbm_Tip_Bens.EOF Then
                        Tip_Bem = M46V999.Check(lrsTbm_Tip_Bens("Dsc_Bem"))
                        Classe_Bem = M46V999.Check(lrsTbm_Tip_Bens("Classe_Bem"))
                        Observacao = M46V999.Check(lrsTbm_Tip_Bens("Observacao"))
                        gTipBem = Tip_Bem
                    End If
                End If
            End If
        Else
            gTipBem = ""
        End If
    Else
        gTipBem = ""
    End If

SAIDA:
    GoSub FechaObj
    Exit Property
FechaObj:
    Set lbdTbm_Tip_Bens = Nothing
    Set lrsTbm_Tip_Bens = Nothing
    Return
End Function


Public Property Get gClasseBem() As String
    Dim lbdTbm_Tip_Bens As ADODB.Connection    'Arquivo: P00Tarifa.
    Dim lMensagem As String     'Mensagem.
    Dim lrsTbm_Tip_Bens As clsYasRecordSet  'Registro : TAB_COBERT.
    Dim lSelect As String

    If val(Classe_Bem) > 0 Then
        gClasseBem = Classe_Bem
    Else
        If Cod_Bem <> 0 Then
            If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, lbdTbm_Tip_Bens, lMensagem) = False Then
                Call M46V999.gpGraLog(2, lMensagem)
                GoTo SAIDA
            End If

            lSelect = " Select top 1 * from Tbm_Tip_Bens WITH (NOLOCK) "
            lSelect = lSelect & " where Cod_Bem  = " & Cod_Bem
            lSelect = lSelect & "   and Cod_tarif  = " & M46V999.gCodProduto
            lSelect = lSelect & "   and Cod_Sequencia <= " & M46V999.gCodSequencia

            '## smrosa
            If M46V999.gCodProduto >= 947 Then
                lSelect = lSelect & "    and Cod_Ramo  = " & M46V999.gCod_Ramo
            End If

            lSelect = lSelect & "   order by Cod_Sequencia desc"

            If M46V999.gfObtRegistro(lbdTbm_Tip_Bens, lSelect, lrsTbm_Tip_Bens, lMensagem) = False Then
                Call M46V999.gpGraLog(2, lMensagem)
                Call M46V999.gpFecha2(lbdTbm_Tip_Bens)
                GoTo SAIDA
            End If
            If Not lrsTbm_Tip_Bens.EOF Then
                Tip_Bem = M46V999.Check(lrsTbm_Tip_Bens("Dsc_Bem"))
                Classe_Bem = M46V999.Check(lrsTbm_Tip_Bens("Classe_Bem"))
                Observacao = M46V999.Check(lrsTbm_Tip_Bens("Observacao"))
                gClasseBem = Classe_Bem
            End If
        End If
    End If

SAIDA:
    GoSub FechaObj
    Exit Property
FechaObj:
    Set lbdTbm_Tip_Bens = Nothing
    Set lrsTbm_Tip_Bens = Nothing
    Return
End Property
Public Property Get gObservacao() As String
    Dim lbdTbm_Tip_Bens As ADODB.Connection    'Arquivo: P00Tarifa.
    Dim lMensagem As String     'Mensagem.
    Dim lrsTbm_Tip_Bens As clsYasRecordSet  'Registro : TAB_COBERT.
    Dim lSelect As String

    If M46V999.gfPreenchido(Observacao) Then
        gObservacao = Observacao
    Else
        If Cod_Bem <> 0 Then
            If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, lbdTbm_Tip_Bens, lMensagem) = False Then
                Call M46V999.gpGraLog(2, lMensagem)
                GoTo SAIDA
            End If

            lSelect = " Select top 1 * from Tbm_Tip_Bens WITH (NOLOCK) "
            lSelect = lSelect & " where Cod_Bem  = " & Cod_Bem
            lSelect = lSelect & "   and Cod_tarif  = " & M46V999.gCodProduto
            lSelect = lSelect & "   and Cod_Sequencia <= " & M46V999.gCodSequencia

            '## smrosa
            If M46V999.gCodProduto >= 947 Then
                lSelect = lSelect & "    and Cod_Ramo  = " & M46V999.gCod_Ramo
            End If

            lSelect = lSelect & "   order by Cod_Sequencia desc "

            If M46V999.gfObtRegistro(lbdTbm_Tip_Bens, lSelect, lrsTbm_Tip_Bens, lMensagem) = False Then
                Call M46V999.gpGraLog(2, lMensagem)
                Call M46V999.gpFecha2(lbdTbm_Tip_Bens)
                GoTo SAIDA
            End If
            If Not lrsTbm_Tip_Bens.EOF Then
                Tip_Bem = M46V999.Check(lrsTbm_Tip_Bens("Dsc_Bem"))
                Classe_Bem = M46V999.Check(lrsTbm_Tip_Bens("Classe_Bem"))
                Observacao = M46V999.Check(lrsTbm_Tip_Bens("Observacao"))
                gObservacao = Observacao
            End If
        End If
    End If

SAIDA:
    GoSub FechaObj
    Exit Property
FechaObj:
    Set lbdTbm_Tip_Bens = Nothing
    Set lrsTbm_Tip_Bens = Nothing
    Return
End Property
' ======================================================================================================================================================

Public Property Get Class_ID() As String
    Class_ID = CStr(nosso_numero) & _
               CStr(Format(Num_Item, "00000")) & _
               CStr(Format(Cod_Seq, "000"))
End Property

Public Property Get gCopiaBem(pNumItem As Long) As stcA46V715B
    Dim lobjstcBem As stcA46V715B
    Set lobjstcBem = New stcA46V715B
    Set lobjstcBem.M46V999 = M46V999

    With lobjstcBem
        .nosso_numero = nosso_numero
        .Num_Ped = .Num_Ped
        .Cod_Org = Cod_Org
        .Cod_Ramo = Cod_Ramo
        .Num_Apol = Num_Apol
        .Num_Endo = Num_Endo
        .Num_Item = pNumItem
        .Dsc_Bem = Dsc_Bem
        .Val_Bem = Val_Bem
        .Cod_Bem = Cod_Bem
        .Cod_Seq = Cod_Seq
    End With
    Set gCopiaBem = lobjstcBem
End Property

Public Property Set ObjPai(pobjstc As stcA46V708B)
    Set ObjStcPai = pobjstc
End Property

Public Property Get Cod_Unicidade() As Long
    If Not ObjStcPai Is Nothing Then
        Cod_Unicidade = ObjStcPai.Cod_Unicidade
    End If
End Property

Public Function getRpt() As stcA46V715Brpt
    Static lrpt As stcA46V715Brpt

    If lrpt Is Nothing Then
        Set lrpt = New stcA46V715Brpt
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
