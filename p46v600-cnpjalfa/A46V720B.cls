VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V720B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ======================================================================================================================================================

' Class Name: stcA46V720B --> Tabela Tab_Ped_Fat_Risco
' Data/Hora: 11/08/2006 - 09:52:48
' Sistema: Multi
' Simony Sales
' ======================================================================================================================================================


' ======================================================================================================================================================
Public nosso_numero As String       'Nosso número.
Public Num_Ped As Long    'Número do pedido.
Public Cod_Org As Integer    'Código do órgão.
Public Cod_Ramo As e_CodRamo    'Código do ramo.
Public Num_Apol As Double     'Sem uso - número da apólice.
Public Num_Endo As Long     'Sem uso - número do endosso.
Public Num_Item As Long     'Número do item.
Public Cod_Fator As Integer   'Código do Fator de Proteção
Public Dsc_Compl_Fator As String          'Complemento do Fator de Proteção
'Public NUM_PI As Long   'Número do PI.
Public NUM_PI As Double   'Número do PI.
Public Desc_Fator As String   'Descricao do Fator de Proteção
' ======================================================================================================================================================
Private ObjStcPai As stcA46V708B
' ======================================================================================================================================================

Public Property Get Class_ID() As String
    Class_ID = CStr(nosso_numero) & _
               CStr(Format(Num_Item, "00000")) & _
               CStr(Format(Cod_Fator, "0000"))
End Property


Public Property Get gCopiaFatores(pNumItem As Long) As stcA46V720B
    Dim lobjstcFator As stcA46V720B
    Set lobjstcFator = New stcA46V720B
    With lobjstcFator
        .nosso_numero = nosso_numero
        .Num_Ped = .Num_Ped
        .Cod_Org = Cod_Org
        .Cod_Ramo = Cod_Ramo
        .Num_Apol = Num_Apol
        .Num_Endo = Num_Endo
        .Num_Item = pNumItem
        .Cod_Fator = Cod_Fator
        .Dsc_Compl_Fator = Dsc_Compl_Fator
        .NUM_PI = NUM_PI
    End With
    Set gCopiaFatores = lobjstcFator
End Property

Public Property Set ObjPai(pobjstc As stcA46V708B)
    Set ObjStcPai = pobjstc
End Property

Public Property Get Cod_Unicidade() As Long
    If Not ObjStcPai Is Nothing Then
        Cod_Unicidade = ObjStcPai.Cod_Unicidade
    End If
End Property

Public Property Get EhDescontoEspecial() As Boolean
    Dim lretorno As Boolean
    '1200539 - Tarifa 1038 e 1039 Junho 2012 - Item 8
    'Retirado na ficha 1200733
    If (Cod_Ramo = e_CodRamo_Residencial) And (Cod_Fator = 7) Then
        If (ObjStcPai.ObjPai.Cod_Produto >= 1022 And ObjStcPai.ObjPai.Cod_Produto <= 1033) Or _
           (ObjStcPai.ObjPai.Cod_Produto >= 1038 And ObjStcPai.ObjPai.Cod_Produto <= 1043) Or _
           (ObjStcPai.ObjPai.Cod_Produto >= 1048) Then
            lretorno = True
        End If
    Else
        lretorno = False
    End If
    EhDescontoEspecial = lretorno
End Property
