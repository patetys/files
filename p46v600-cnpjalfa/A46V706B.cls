VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V706B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================

' Class Name: stcA46V706B --> Tabela Tab_Ped_Coss
' Data/Hora: 31/08/2005 - 11:53:36
' Sistema: Multi
' Thais Juren Gomes

' ======================================================================================================================================================


' ======================================================================================================================================================
Public nosso_numero As String           'Nosso número.
Public Cod_Coss As Long         'Código da congenere.
Public Num_Ped As Long        'Número do pedido.
Public Cod_Org As Long        'Código do órgão.
Public Cod_Ramo As e_CodRamo    'Código do Ramo.
Public Num_Apol As Double       'Sem uso - número da apólice.
Public Num_Endo As Long         'Sem uso - número do endosso.
Public Per_Coss As Double       'Percentual.
Public NUM_PI As Double     'Número do PI.
Public Nom_Seguradora As String             'Nome da seguradora (não é campo do banco de dados).
' ======================================================================================================================================================
Private ObjStcPai As stcA46V702B
' ======================================================================================================================================================
Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "stcA46V706B"
End Property

Public Property Get Class_ID() As String

    Class_ID = CStr(nosso_numero) & _
               CStr(Format(Cod_Coss, "0000000"))
End Property
Public Property Get gNumOrdem() As String
    gNumOrdem = ""
End Property
Public Property Get gNom_Seguradora() As String
    Dim lMensagem As String
    Dim lNomeSeg As String
    Dim lNomAbrev As String
    If Not M46V999.gfPreenchido(Nom_Seguradora) Then
        If Not M46V999.gfNomSeguradora(Me.Cod_Coss, lNomeSeg, lMensagem, lNomAbrev) Then
            Call M46V999.gpGraLog(1, lMensagem)
            Exit Property
        End If
        Me.Nom_Seguradora = lNomAbrev
    End If
    gNom_Seguradora = Me.Nom_Seguradora
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
