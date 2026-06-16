VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V705B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================

' Class Name: stcA46V705B --> Tabela Tab_Ped_Corr
' Data/Hora: 31/08/2005 - 11:53:11
' Sistema: Multi
' Thais Juren Gomes

' ======================================================================================================================================================


' ======================================================================================================================================================
Public nosso_numero As String   'Nosso número.
Public Cod_Corr As Long    'Código do corretor.
Public Num_Ped As Long    'Número do pedido.
Public Cod_Org As Long    'Código do órgão.
Public Cod_Ramo As e_CodRamo    'Código do Ramo.
Public Num_Apol As Double    'Sem uso - número da apólice.
Public Num_Endo As Long    'Sem uso - número do endosso.
Public Per_Corr As Double    'Percentual.
Public Num_Ordem_Corr As Long       'Número da ordem.
Public NUM_PI As Double    'Número do PI.
Public mNumSusep As String


Public ComparaCorretor As String
' ======================================================================================================================================================
Private ObjStcPai As stcA46V702B
' ======================================================================================================================================================
Private mNomeCorretor As String
Private mCodCorretor As Long

Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "A46V705B"
End Property

Public Property Get Class_ID() As String
    Class_ID = CStr(nosso_numero) & _
               CStr(Format(Cod_Corr, "0000000")) & CStr(Format(Cod_Ramo, "000"))
End Property

Public Property Get gNom_Corretor() As String
    Dim lSaida As Variant
    Dim lMensagem As String
    If mCodCorretor <> Cod_Corr Or mNomeCorretor = "" Then
        If M46V999.gfObtCorretor(1, Me.Cod_Corr, lSaida, lMensagem) = True Then
            mNomeCorretor = lSaida(1, 3)
            mNumSusep = lSaida(1, 4)
        Else    'Quando não localizar o corretor ou estiver inavivo.
            mCodCorretor = 0
            Cod_Corr = 0
        End If
        mCodCorretor = Cod_Corr
    End If
    gNom_Corretor = mNomeCorretor
End Property

Public Property Set ObjPai(pobjstc As stcA46V702B)
    Set ObjStcPai = pobjstc
End Property

Public Property Get Cod_Unicidade() As Long
    If Not ObjStcPai Is Nothing Then
        Cod_Unicidade = ObjStcPai.Cod_Unicidade
    End If
End Property

Public Property Set M46V999(pClsM46V999 As clsM46V999)
    Set mClsM46V999 = pClsM46V999
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

Private Sub Class_Initialize()
    Set mClsM46V999 = Nothing
End Sub

Private Sub Class_Terminate()
    Set mClsM46V999 = Nothing
End Sub
