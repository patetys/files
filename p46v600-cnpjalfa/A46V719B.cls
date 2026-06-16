VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V719B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================

' Class Name: stcA46V719B --> Tabela Tab_Ped_Auto
' Data/Hora: 08/08/2006 - 15:29:14
' Sistema: Multi
' Simony Sales
' ======================================================================================================================================================


' ======================================================================================================================================================
Public nosso_numero As String           'Nosso número.
Public Num_Ped As Long        'Número do pedido.
Public Cod_Org As Long        'Código do órgão.
Public Cod_Ramo As e_CodRamo    'Código do ramo.
Public Num_Apol As Double       'Sem uso - número da apólice.
Public Num_Endo As Long         'Sem uso - número do endosso.
Public Val_Is_Auto As Double
Public Val_Is_Rc As Double
Public Val_Is_App As Double
Public Val_Pr_Auto As Double
Public Val_Pr_Rc As Double
Public Val_Pr_App As Double
Public NUM_PI As Double     'PI.

' ======================================================================================================================================================
Private ObjStcPai As stcA46V702B
' ======================================================================================================================================================

Public Property Get Class_ID() As String

    Class_ID = CStr(nosso_numero)
End Property

Public Property Set ObjPai(pobjstc As stcA46V702B)
    Set ObjStcPai = pobjstc
End Property

Public Property Get Cod_Unicidade() As Long
    If Not ObjStcPai Is Nothing Then
        Cod_Unicidade = ObjStcPai.Cod_Unicidade
    End If
End Property
