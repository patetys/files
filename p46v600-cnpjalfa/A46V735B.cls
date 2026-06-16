VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V735B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: stcA46V735B - Classe de Beneficiario, Inquilino e Proprietario
' Data/Hora: 05/11/2013 - 10:28:32
' Sistema: Multi
' Simony Sales
' ======================================================================================================================================================
Public nosso_numero As String    'PK
Public Num_Item As Long         'PK                 'N�mero do item.
Public NUM_PI As Double
Public Num_Apol As Double     'Sem uso - n�mero da ap�lice.
Public Num_Endo As Long     'Sem uso - n�mero do endosso

Public Ped_Loc_Pessoa_ID As Long
Public Tip_Loc_Pessoa As Integer
Public Nom_Pessoa As String
Public Tip_Pessoa As Integer
Public Num_CPF_CNPJ As String
Public Tip_Cobert As Integer
Public Per_Participacao As Double
Public Tip_Doc As Integer
Public Num_Doc As String

' ======================================================================================================================================================
Private ObjStcPai As stcA46V708B

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
               CStr(Format(Ped_Loc_Pessoa_ID, "00000"))
End Property

Public Property Get gCopiaPessoa(pNum_Item As Long) As stcA46V735B
    Dim lObjstcObs As stcA46V735B

    Set lObjstcObs = New stcA46V735B
    With lObjstcObs
        .nosso_numero = nosso_numero
        .Num_Apol = Num_Apol
        .NUM_PI = NUM_PI
        .Num_Item = pNum_Item
        '.Ped_Loc_Pessoa_ID = Ped_Loc_Pessoa_ID
        .Tip_Loc_Pessoa = Tip_Loc_Pessoa
        .Nom_Pessoa = Nom_Pessoa
        .Tip_Pessoa = Tip_Pessoa
        .Num_CPF_CNPJ = Num_CPF_CNPJ
        .Tip_Cobert = Tip_Cobert
        .Per_Participacao = Per_Participacao
        .Tip_Doc = Tip_Doc
        .Num_Doc = Num_Doc
    End With

    Set gCopiaPessoa = lObjstcObs
End Property

Public Property Get M46V999() As clsM46V999
    If mClsM46V999 Is Nothing Then
        Err.Raise 91, Me.Class_Name, "Objeto M46V999 n�o atribu�do para a classe (" & Me.Class_Name & ")"
        '***Esta Classe DEVE utilizar uma instancia da classe clsM46V999 j�
        '   iniciada no escopo onde esta classe foi instanciada.
        'Atribua a o objeto (clsM46V999 ja instanciado) no momento que eu inicializou esta classe (atraves do Property SET)!!!
    End If

Retorno:
    Set M46V999 = mClsM46V999
End Property

Public Property Get Class_Name() As String
    Class_Name = "A46V735B"
End Property
Public Property Set M46V999(pClsM46V999 As clsM46V999)
    Set mClsM46V999 = pClsM46V999
End Property

Public Property Set ObjPai(pobjstc As stcA46V708B)
    Set ObjStcPai = pobjstc
End Property

Public Property Get Cod_Unicidade() As Long
    If Not ObjStcPai Is Nothing Then
        Cod_Unicidade = ObjStcPai.Cod_Unicidade
    End If
End Property

