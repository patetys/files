VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V716B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: stcA46V716B -- Tab_Ped_Loc_Seguro
' Data/Hora: 26/09/2005 - 11:34:01
' Sistema: Multi
' Thais Juren Gomes
' ======================================================================================================================================================

Public nosso_numero As String   'Nosso número.
Public Num_Item As Long    'Número do item.
Public Cod_Org As Long    'Código do orgão.
Public Cod_Ramo As Long    'Código do Ramo.
Public Num_Ped As Long    'Número do pedido.
Public Num_Apol As Double    'Número da apólice.
Public Num_Endo As Long    'Número do endosso.
Public Dsc_Outro_Seguro As String       'Descrição - não utilizado.
Public Cod_Seguradora As Long       'Código da seguradora.
Public Num_Ordem As Long  'Número da ordem.
Public Num_Apol_Ant As String   'Número da apólice anterior
Public Num_Item_Ant As Long     'Número do item anterior
Public Ind_Sinistro As Byte         'e_SimNao         'Indicador de sinistro. ***MOSILVA Verificar!!!
' ======================================================================================================================================================
Private ObjStcPai As stcA46V708B
' ======================================================================================================================================================

Public Property Get Class_ID() As String
    Class_ID = CStr(nosso_numero) & _
               CStr(Format(Num_Item, "00000")) & _
               Format(Me.Cod_Seguradora, "0000")
End Property

Public Property Get gCopiaSeguro(pNumItem As Long) As stcA46V716B
    Dim lobjstcSeguro As stcA46V716B
    Set lobjstcSeguro = New stcA46V716B
    With lobjstcSeguro
        .nosso_numero = nosso_numero
        .Num_Ped = .Num_Ped
        .Cod_Org = Cod_Org
        .Cod_Ramo = Cod_Ramo
        .Num_Apol = Num_Apol
        .Num_Endo = Num_Endo
        .Num_Item = pNumItem
        .Dsc_Outro_Seguro = Dsc_Outro_Seguro
        .Cod_Seguradora = Cod_Seguradora
        .Num_Ordem = Num_Ordem
        .Num_Apol_Ant = Num_Apol_Ant
        .Num_Item_Ant = Num_Item_Ant
        .Ind_Sinistro = Ind_Sinistro
    End With
    Set gCopiaSeguro = lobjstcSeguro
End Property

Public Property Set ObjPai(pobjstc As stcA46V708B)
    Set ObjStcPai = pobjstc
End Property

Public Property Get Cod_Unicidade() As Long
    If Not ObjStcPai Is Nothing Then
        Cod_Unicidade = ObjStcPai.Cod_Unicidade
    End If
End Property
