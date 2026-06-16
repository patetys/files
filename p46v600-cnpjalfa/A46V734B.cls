VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V734B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'--Tab_Ass24h

'**tarifa
Public Cod_Tarif As Integer             'Código do produto.
Public Cod_Sequencia As Integer
'--
Public Cod_Ramo As e_CodRamo
Public Cod_Modalidade As Integer
'--
'**Opcao
Public Cod_Clau As Long
'Residencial
'400 - Opção A
'401 – Opção B
'402 – Opção C
'
'Empresarial
'400 - Opção A
'401 – Opção B
'402 – Opção C
'
'Condomínio
'400 - Opção A
'--
Public Premio As Double

Private mClsM46V999 As clsM46V999
Private Sub Class_Initialize()
    Set mClsM46V999 = Nothing
End Sub

Private Sub Class_Terminate()
    Set mClsM46V999 = Nothing
End Sub
Public Property Get Class_Name() As String
    Class_Name = "A46V734B"
End Property

Public Property Get Class_ID() As String
    Class_ID = CStr(Cod_Tarif) & _
               CStr(Cod_Sequencia) & _
               CStr(Cod_Ramo) & _
               CStr(Cod_Modalidade) & _
               CStr(Cod_Clau)
End Property

Public Property Get Descricao() As String
    Dim lStcClau As stcA46V703B
    Static lsRetorno As String
    Static lsCodClau As Long
    Static lsCodRamo As e_CodRamo

    If Me.Cod_Clau <> lsCodClau And _
       Me.Cod_Ramo <> lsCodRamo Then

        lsCodClau = Me.Cod_Clau
        lsCodRamo = Me.Cod_Ramo

        Set lStcClau = New stcA46V703B
        Set lStcClau.M46V999 = M46V999

        lStcClau.Cod_Clau = lsCodClau
        lStcClau.Cod_Ramo = lsCodRamo

        lsRetorno = Trim(lStcClau.gDescClausula)

        Set lStcClau = Nothing
    End If

    Descricao = lsRetorno

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



