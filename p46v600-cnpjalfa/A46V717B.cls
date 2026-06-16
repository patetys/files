VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V717B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: stcA46V717B --Tab_Ped_Insp
' Data/Hora: 26/09/2005 - 17:03:11
' Sistema: Multi
' Thais Juren Gomes
' ======================================================================================================================================================
Public Num_Item As Long    'Número do item
Public nosso_numero As String   'Nosso número
Public Cod_Org As Long    'Código do órgão
Public Cod_Ramo As Long    'Código do Ramo
Public Num_Ped As Long    'Número do pedido
Public Num_Insp As Long    'Número da inspeção
Public Dat_Cadastro As Long     'Data do cadastro (yyyymmdd)
Public Pess_Cont As String    'Contato.
Public Num_Tel As String    'Número do telefone
Public Num_Logr_Risco As Long       'Número do logradouro do risco.
'Public NUM_PI As Long
Public NUM_PI As Double
Public PESS_CONTATO_2 As String
Public PESS_CONTATO_3 As String
Public NUM_TEL_2 As String
Public NUM_TEL_3 As String
Public DSC_EMAIL As String
Public DSC_PONTO_REF As String
' ======================================================================================================================================================
Private ObjStcPai As stcA46V708B
' ======================================================================================================================================================

Public Property Get Class_ID() As String
    Class_ID = CStr(nosso_numero) & _
               Format(Num_Item, "00000")
End Property

Public Property Get gCopiaInsp(pNumItem As Long) As stcA46V717B
    Dim lObjstcInsp As stcA46V717B
    Set lObjstcInsp = New stcA46V717B
    With lObjstcInsp
        .nosso_numero = nosso_numero
        .Num_Ped = Num_Ped
        .Cod_Org = Cod_Org
        .Cod_Ramo = Cod_Ramo
        .Num_Item = pNumItem
        .Num_Insp = Num_Insp
        .Dat_Cadastro = Dat_Cadastro
        .Pess_Cont = Pess_Cont
        .Num_Tel = Num_Tel
        .Num_Logr_Risco = Num_Logr_Risco
    End With
    Set gCopiaInsp = lObjstcInsp
End Property

Public Property Set ObjPai(pobjstc As stcA46V708B)
    Set ObjStcPai = pobjstc
End Property

Public Property Get Cod_Unicidade() As Long
    If Not ObjStcPai Is Nothing Then
        Cod_Unicidade = ObjStcPai.Cod_Unicidade
    End If
End Property
