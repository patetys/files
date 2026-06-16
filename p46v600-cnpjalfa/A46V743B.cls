VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V743B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ======================================================================================================================================================
' Class Name: stcA46V743B -- Tab_Ctrl_Emis
' Data/Hora: 19/09/2005 - 10:20:46
' Sistema: Multi
' Thais Juren Gomes
' ======================================================================================================================================================
Public nosso_numero As String           'PK               'Nosso número
Public NUM_SEQ                  As Double                   'NUM SEQ
'Public NUM_PI                   As Long                     'NUM PI
Public NUM_PI                   As Double                     'NUM PI
Public NUM_APOL_MOV             As Double                   'NUM APOL MOV
Public NUM_COMPL_MOV            As Long                     'NUM COMPL MOV
Public REF_MOV                  As Long                     'REF MOV
Public TIPO_MOV                 As String                   'TIPO MOV
Public VL_PREMIO_LIQ_MOV        As Double                   'VL PREMIO LIQ MOV
Public QTD_ITEM_MOV             As String                   'QTD ITEM MOV
Public DAT_GRAV                 As Long                     'DAT GRAV
Public Cod_User                 As String                   'COD USER
Public DIG_BLOQUETO             As Byte
Private mDigito As String
Private mBloquetoDigito As Long

Public ColecaoCanalEmiss As Collection
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

