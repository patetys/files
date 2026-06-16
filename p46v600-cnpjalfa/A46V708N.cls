VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V708N"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Num_Item As Long         'PK                 'Número do item.
Public COD_CLASSE_NORMA As String
Public COD_ERR_NORMA As String
Public COD_STAT_IRREG As String
Public DSC_ERR_NORMA_INT As String
Public DSC_ERR_NORMA_EXT As String


 
Public Property Get Class_ID() As String
    Class_ID = CStr(Format(Num_Item, "00000")) & _
               CStr(Format(COD_ERR_NORMA, "00000")) & _
               CStr(Format(COD_STAT_IRREG, "00000"))
End Property
