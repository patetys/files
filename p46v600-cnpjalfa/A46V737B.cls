VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V737B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: stcA46V708B --> Tabela Tab_Ped_Loc_Adic
' Data/Hora: 03/01/2014 - 14:00:00
' Sistema: Multi
' Marcio Oliveira da Silva
' ======================================================================================================================================================

Public nosso_numero         As String 'PK               'Nosso número.
Public Num_Item             As Long 'PK                 'Número do item.
Public NUM_SEQ              As Long 'PK                 'Número da Sequencia do Local de Risco.
Public Cod_Tip_Logr_Risco   As String                   'Tipo de logradouro - Risco.
Public Nom_Logr_Risco       As String                   'Nome do logradouro - Risco.
Public Dsc_Compl_Risco      As String                   'Complemento do logradouro - Risco.
Public Nom_Bairro_Risco     As String                   'Bairro - Risco.
Public Nom_Cid_Risco        As String                   'Cidade - Risco.
Public Sig_UF_Risco         As String                   'UF - Risco.
Public Num_Cep_Risco        As Long                     'CEP - Risco.

Private ObjStcPai           As stcA46V708B

Public Property Get Class_Name() As String
    Class_Name = "A46V708B"
End Property

Public Property Get Class_ID() As String
    Class_ID = CStr(nosso_numero) & _
               CStr(Format(Num_Item, "00000")) & _
               CStr(Format(NUM_SEQ, "00000"))
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

Private Sub Class_Initialize()
    Set ObjStcPai = Nothing
End Sub

Private Sub Class_Terminate()
    Set ObjStcPai = Nothing
End Sub
