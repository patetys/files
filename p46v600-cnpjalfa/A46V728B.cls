VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V728B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: class1  --> Tabela Tab_Ped_Loc_LMI
' Data/Hora: 13/12/2007 - 11:00:41
' Sistema: Multi
' Adriano Tirabasso
' ======================================================================================================================================================

Public nosso_numero As String           'Nosso número.
Public Num_Item As Long         'PK                 'Número do item.
Public Num_Ped As Long         'PK                 'Número do pedido.
Public Cod_Org As Integer     'Código do órgão.
Public Cod_Ramo As e_CodRamo    'Código do ramo.
Public Num_Apol As Double         'Número da apólice.
Public Num_Endo As Long         'Número do endosso.
Public Nom_Logr As String       'Nome do logradouro.
Public Nom_Bairro As String         'Nome do Bairro.
Public Nom_Cid As String      'Cidade.
Public Sig_UF As String     'UF.
Public Num_CEP As String      'CEP.
Public NUM_PI As Double     'Número do PI
Public NUM_SEQ As String      'Número da Sequência

' ======================================================================================================================================================
Private ObjStcPai As stcA46V708B
' ======================================================================================================================================================

Public Property Get Class_ID() As String
    Class_ID = CStr(nosso_numero) & _
               CStr(Format(Num_Item, "00000")) & _
               CStr(Format(NUM_SEQ, "000"))
End Property


Public Property Get gCopiaLMI(pNum_Item As Long) As stcA46V728B
    Dim lobjstcLMI As stcA46V728B

    Set lobjstcLMI = New stcA46V728B
    With lobjstcLMI
        .nosso_numero = Me.nosso_numero
        .Num_Item = pNum_Item
        .Num_Ped = Num_Ped
        .Cod_Org = Cod_Org
        .Cod_Ramo = Cod_Ramo
        .Num_Apol = Num_Apol
        .Num_Endo = Num_Endo
        .Nom_Logr = Nom_Logr
        .Nom_Bairro = Nom_Bairro
        .Nom_Cid = Nom_Cid
        .Sig_UF = Sig_UF
        .Num_CEP = Num_CEP
        .NUM_PI = NUM_PI
        .NUM_SEQ = NUM_SEQ
    End With
    Set gCopiaLMI = lobjstcLMI
End Property
Public Property Get gCEP_Formatado() As String
    Dim lAux As String
    If Num_CEP <> 0 Then
        lAux = Format(Num_CEP, "00000000")
        gCEP_Formatado = Mid$(lAux, 1, 5) & "-" & Mid$(lAux, 6, 3)
    Else
        gCEP_Formatado = " "
    End If

End Property

Public Property Set ObjPai(pobjstc As stcA46V708B)
    Set ObjStcPai = pobjstc
End Property

Public Property Get Cod_Unicidade() As Long
    If Not ObjStcPai Is Nothing Then
        Cod_Unicidade = ObjStcPai.Cod_Unicidade
    End If
End Property
