VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsYasField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mName As String
Private mValue As Variant
Private mType As DataTypeEnum

Public Property Get Name() As String
    Name = mName
End Property

Public Property Let Name(ByVal pNome As String)
    mName = pNome
End Property

Public Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0
'Propriedade DEFAULT!!!
    Value = mValue
End Property

Public Property Let Value(ByVal pValor As Variant)
    mValue = pValor
End Property

Public Property Get Type_() As DataTypeEnum
    Type_ = mType
End Property

Public Property Let Type_(ByVal pType As DataTypeEnum)
    mType = pType
End Property
