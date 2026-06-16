VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V738B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name:  - Tab_Ped_Cobert_Franquia
' Data/Hora: 22/07/2014 - 10:01:32
' Sistema: Multi
' Rodrigo Berçot
' ======================================================================================================================================================
Public Cod_Ramo As Integer    'PK
Public Cod_Modalidade As Integer
Public Cod_Tarif As Integer     'Sem uso - número da apólice.
Public Cod_Cobert As Integer     'Sem uso - número do endosso

Public Cod_Evento As Integer
Public Num_Ped As Long
Public Num_Apol As Double
Public Num_Item As Integer
Public Num_Endosso As Long
Public nosso_numero As String
'Public NUM_PI As Long
Public NUM_PI As Double
Public Id_Franquia_Avancada As Integer
Public Val_Min_Franq As Double
Public Per_Franq As Double
Public Texto_Franquia As String
Public Texto_Observacoes As String
Public Multiplo_Franquia As Integer
' ======================================================================================================================================================
Private ObjStcPai As stcA46V704B

Private mClsM46V999 As clsM46V999
Private Sub Class_Initialize()
    Set mClsM46V999 = Nothing
End Sub

Private Sub Class_Terminate()
    Set mClsM46V999 = Nothing
End Sub

' ======================================================================================================================================================
Public Property Get Class_ID() As String
    Class_ID = CStr(nosso_numero) & _
               CStr(Format(Num_Item, "00000")) & _
               CStr(Format(Cod_Cobert, "0000")) & _
               CStr(Format(Cod_Modalidade, "00000")) & _
               CStr(Format(Multiplo_Franquia, "00"))
End Property

Public Property Get gCopiaTab_Ped_Franquia(pNum_Item As Long) As stcA46V738B
    Dim lObjstcObs As stcA46V738B

    Set lObjstcObs = New stcA46V738B
    With lObjstcObs
        .Cod_Ramo = Cod_Ramo
        .Cod_Modalidade = Cod_Modalidade
        .Cod_Tarif = Cod_Tarif
        .Cod_Cobert = Cod_Cobert
        .Cod_Evento = Cod_Evento
        .Num_Ped = Num_Ped
        .Num_Apol = Num_Apol
        .Num_Endosso = Num_Endosso
        .nosso_numero = nosso_numero
        .Val_Min_Franq = Val_Min_Franq
        .Per_Franq = Per_Franq
        .Texto_Franquia = Texto_Franquia
        .Texto_Observacoes = Texto_Observacoes
        .Multiplo_Franquia = Multiplo_Franquia
    End With

    Set gCopiaTab_Ped_Franquia = lObjstcObs
End Property

Public Property Get M46V999() As clsM46V999
    If mClsM46V999 Is Nothing Then
        Err.Raise 91, Me.Class_Name, "Objeto M46V999 não atribuído para a classe (" & Me.Class_Name & ")"
        '***Esta Classe DEVE utilizar uma instancia da classe clsM46V999 já
        '   iniciada no escopo onde esta classe foi instanciada.
        'Atribua a o objeto (clsM46V999 ja instanciado) no momento que eu inicializou esta classe (atraves do Property SET)!!!
    End If

Retorno:
    Set M46V999 = mClsM46V999
End Property

Public Property Get Class_Name() As String
    Class_Name = "A46V738B"
End Property
Public Property Set M46V999(pClsM46V999 As clsM46V999)
    Set mClsM46V999 = pClsM46V999
End Property

Public Property Set ObjPai(pobjstc As stcA46V704B)
    Set ObjStcPai = pobjstc
End Property

Public Property Get Cod_Unicidade() As Long
    If Not ObjStcPai Is Nothing Then
        Cod_Unicidade = ObjStcPai.Cod_Unicidade
    End If
End Property



