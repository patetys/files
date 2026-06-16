VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsYasRecordSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private registros As Collection

Private regAtual As Long
Private mFieldsCount As Long

Private Sub Class_Initialize()
    Call Inicializar
End Sub

Private Sub Class_Terminate()
    Set registros = Nothing
End Sub

Private Sub Inicializar()
    Set registros = New Collection
    regAtual = 0
End Sub

Public Sub LoadRecordSet(ByRef pRs As ADODB.Recordset)
    Dim lCampos As Collection
    Dim lCampo As clsYasField
    Dim i As Integer
    Dim lFielNames As Variant

    Call Inicializar

    If Not pRs Is Nothing Then
        lFielNames = mfGeraFieldNames(pRs)
        Do While pRs.EOF = False
            Set lCampos = New Collection
            mFieldsCount = pRs.Fields.Count
            For i = 0 To (pRs.Fields.Count - 1)
                Set lCampo = New clsYasField
                With lCampo
                    '                    Debug.Print "Campo(" & I & ") "; lFielNames(I) & "<==>" & pRs.Fields(I).Name
                    .Name = lFielNames(i)    'pRs.Fields(I).Name
                    .Value = pRs.Fields(i).Value
                    .Type_ = pRs.Fields(i).Type
                End With
                lCampos.Add lCampo, lCampo.Name
            Next i
            registros.Add lCampos
            Set lCampos = Nothing
            pRs.MoveNext
        Loop
        pRs.Close
        Set pRs = Nothing
    End If

    If registros.Count > 0 Then
        regAtual = 1
    End If
End Sub

Private Function mfGeraFieldNames(ByRef pRs As ADODB.Recordset) As Variant
    Dim lStrAux As String
    Dim lArrFieldNames() As String
    Dim lColecaoDummy As Collection
    Dim i As Integer

    If Not pRs Is Nothing Then
        If pRs.Fields.Count > 0 Then
            Set lColecaoDummy = New Collection
            ReDim lArrFieldNames(pRs.Fields.Count - 1) As String
            For i = 0 To (pRs.Fields.Count - 1)
                lStrAux = pRs.Fields(i).Name
                lArrFieldNames(i) = mfNovoFiledName(lColecaoDummy, lStrAux)
            Next i
        End If
    End If
    mfGeraFieldNames = lArrFieldNames
End Function

Private Function mfNovoFiledName(ByRef pColecao As Collection, ByVal pName As String) As String
    Dim lSufix As String
    Dim i As Integer
    Dim lretorno As String

    Err.Clear
    On Error Resume Next
    lretorno = pName
    Err = -1
    Do While (Err <> 0 Or i > 10)
        lretorno = pName & lSufix
        Err.Clear
        pColecao.Add pName, lretorno
        i = i + 1
        lSufix = "(" & Format(i, "000") & ")"
    Loop
    On Error GoTo 0
SAIDA:
    mfNovoFiledName = lretorno
End Function

Public Property Get EOF() As Boolean
    EOF = (regAtual = 0 Or regAtual > registros.Count)
End Property

Public Property Get BOF() As Boolean
    BOF = (regAtual = 0)
End Property

Public Sub MoveNext()
    regAtual = regAtual + 1
End Sub

Public Sub MoveFirst()
    regAtual = IIf(registros.Count > 0, 1, 0)
End Sub

Public Sub MoveLast()
    regAtual = registros.Count
End Sub

Public Property Get Fields(Optional ByVal index) As Object
Attribute Fields.VB_UserMemId = 0

'Propriedade DEFAULT!!!
    If (IsMissing(index)) Then
        'Retorna a colecao de campos (clsYasField)
        Set Fields = registros(regAtual)
    Else
        If (Len(index) = 0) Then
            Set Fields = registros(regAtual)
        Else
            'Retorna um Campo (clsYasField)
            If index = 0 Then
                index = 1    '*** Diferentemente do RecosrSeto Campo inicia no index 1
            End If
            Err.Clear
            On Error Resume Next
            Set Fields = registros(regAtual).Item(index)
            If Err = 5 Then
                Err = 3265
            End If
        End If
    End If
End Property

Public Property Get RecordCount() As Long
    RecordCount = registros.Count
End Property

Public Property Get AbsolutePosition() As Long
    AbsolutePosition = regAtual
End Property
Public Property Get FieldsCount() As Long
    FieldsCount = mFieldsCount
End Property
Public Property Get Colecao() As Collection
    Set Colecao = registros.Item(1)
End Property
