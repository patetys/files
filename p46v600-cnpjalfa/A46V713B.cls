VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V713B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: stcA46V713B -- Tab_Ped_Venc
' Data/Hora: 20/09/2005 - 15:56:32
' Sistema: Multi
' Thais Juren Gomes
' ======================================================================================================================================================
Public nosso_numero As String   'Nosso número.
Public Cod_Org As Long    'Código do orgao.
Public Cod_Ramo As Long    'Código do ramo.
Public Num_Ped As Long    'Número do pedido.
Public Num_Apol As Double    'Número da apólice.
Public Num_Endo As Long    'Número do endosso.
Public Num_Parc As Long    'Número da parcela.
Public DAT_VENC As Long    'Data do vencimento da parcela.
' ======================================================================================================================================================
Private ObjStcPai As stcA46V702B

Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "A46V713B"
End Property

Public Property Get Class_ID() As String
    Class_ID = CStr(nosso_numero) & _
               Format(CStr(Num_Parc), "00")
End Property

Public Property Get gDataVenc_Formatada() As String
    gDataVenc_Formatada = M46V999.gfForData(2, Me.DAT_VENC)
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
