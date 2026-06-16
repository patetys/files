VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V726B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: stcA46V726B -- Texto para Especificação
' Data/Hora: 05/06/2007 - 10:28:32
' Sistema: Multi
' Simony Sales
' ======================================================================================================================================================
Public nosso_numero As String    'PK
Public Cod_Espec As Long
Public Cod_Org As Integer    'Código do órgão.
Public Cod_Ramo As e_CodRamo    'Código do ramo.
Public Num_Ped As Long    'Número do pedido.
Public Num_Apol As Double     'Sem uso - número da apólice.
Public Num_Endo As Long     'Sem uso - número do endosso
Public Num_pag As Integer
Public Qtd_vias As Integer
Public Espec_lin1 As String
Public Espec_lin2 As String
Public Espec_lin3 As String
Public Espec_lin4 As String
Public Espec_lin5 As String
Public Espec_lin6 As String
Public Espec_lin7 As String
Public Espec_lin8 As String
Public Espec_lin9 As String
Public Espec_lin10 As String
Public Espec_lin11 As String
Public Espec_lin12 As String
Public Espec_lin13 As String
Public Espec_lin14 As String
Public Espec_lin15 As String
Public Espec_lin16 As String
Public Espec_lin17 As String
Public Espec_lin18 As String
Public Espec_lin19 As String
Public Espec_lin20 As String
Public NUM_PI As Double
' ======================================================================================================================================================
Private ObjStcPai As stcA46V702B
' ======================================================================================================================================================
Public Property Get Class_ID() As String
    Class_ID = CStr(nosso_numero) & _
               CStr(Format(Cod_Espec, "00000")) & _
               CStr(Format(Num_pag, "00000"))
End Property

Public Property Get gCopiaTextoEspecificacao() As stcA46V726B
    Dim lObjstcObs As stcA46V726B

    Set lObjstcObs = New stcA46V726B
    With lObjstcObs
        .nosso_numero = nosso_numero
        .Cod_Org = Cod_Org
        .Cod_Ramo = Cod_Ramo
        .Num_Ped = Num_Ped
        .Num_Apol = Num_Apol
        .Num_Endo = Num_Endo
        .Num_pag = Num_pag
        .Qtd_vias = Qtd_vias
        .Espec_lin1 = Espec_lin1
        .Espec_lin2 = Espec_lin2
        .Espec_lin3 = Espec_lin3
        .Espec_lin4 = Espec_lin4
        .Espec_lin5 = Espec_lin5
        .Espec_lin6 = Espec_lin6
        .Espec_lin7 = Espec_lin7
        .Espec_lin8 = Espec_lin8
        .Espec_lin9 = Espec_lin9
        .Espec_lin10 = Espec_lin10
        .Espec_lin11 = Espec_lin11
        .Espec_lin12 = Espec_lin12
        .Espec_lin13 = Espec_lin13
        .Espec_lin14 = Espec_lin14
        .Espec_lin15 = Espec_lin15
        .Espec_lin16 = Espec_lin16
        .Espec_lin17 = Espec_lin17
        .Espec_lin18 = Espec_lin18
        .Espec_lin19 = Espec_lin19
        .Espec_lin20 = Espec_lin20
        .NUM_PI = NUM_PI
    End With
    Set gCopiaTextoEspecificacao = lObjstcObs
End Property

Public Property Set ObjPai(pobjstc As stcA46V702B)
    Set ObjStcPai = pobjstc
End Property

Public Property Get Cod_Unicidade() As Long
    If Not ObjStcPai Is Nothing Then
        Cod_Unicidade = ObjStcPai.Cod_Unicidade
    End If
End Property
