VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V736B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: stcA46V736B - Telefones
' Data/Hora: 25/11/2013 - 10:28:32
' Sistema: Multi
' Simony Sales
' ======================================================================================================================================================
Public nosso_numero As String    'PK
Public NUM_PI As Double
Public Num_Apol As Double     'Sem uso - número da apólice.
Public Num_Endo As Long     'Sem uso - número do endosso

Public Telefone_ID As Long
Public Tip_Telefone As Integer
Public DDD_Telefone As Integer
Public Numero_Telefone As Double
Public Ramal_Telefone As Long
Public Obs_Telefone As String
Public Ind_Principal As Integer

' ======================================================================================================================================================
Private ObjStcPai As stcA46V702B

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
               CStr(Format(Telefone_ID, "0000000000"))
End Property

Public Property Get gCopiaTelefone(pNum_Item As Long) As stcA46V736B
    Dim lObjstcObs As stcA46V736B

    Set lObjstcObs = New stcA46V736B
    With lObjstcObs
        .nosso_numero = nosso_numero
        .Num_Apol = Num_Apol
        .NUM_PI = NUM_PI
        .Telefone_ID = Telefone_ID
        .Tip_Telefone = Tip_Telefone
        .DDD_Telefone = DDD_Telefone
        .Numero_Telefone = Numero_Telefone
        .Ramal_Telefone = Ramal_Telefone
        .Obs_Telefone = Obs_Telefone
        .Ind_Principal = Ind_Principal
    End With

    Set gCopiaTelefone = lObjstcObs
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
    Class_Name = "A46V736B"
End Property
Public Property Set M46V999(pClsM46V999 As clsM46V999)
    Set mClsM46V999 = pClsM46V999
End Property

Public Property Set ObjPai(pobjstc As stcA46V702B)
    Set ObjStcPai = pobjstc
End Property

Public Property Get Cod_Unicidade() As Long
    If Not ObjStcPai Is Nothing Then
        Cod_Unicidade = ObjStcPai.Cod_Unicidade
    End If
End Property

