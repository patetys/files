VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V707B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================

' Class Name: stcA46V707B --> Tabela Tab_Ped_Ender
' Data/Hora: 31/08/2005 - 11:54:14
' Sistema: Multi
' Thais Juren Gomes

' ======================================================================================================================================================


' ======================================================================================================================================================
Public nosso_numero As String           'Nosso número.
Public Num_Ped As Long        'Número do pedido.
Public Cod_Org As Long        'Código do órgão.
Public Cod_Ramo As e_CodRamo    'Código do ramo.
Public Num_Apol As Double       'Sem uso - número da apólice.
Public Num_Endo As Long         'Sem uso - número do endosso.
Public Cod_Tip_Logr As String           'Tipo do logradouro.
Public Nom_Logr As String       'Logradouro.
Public Dsc_Compl_Logr As String             'Complemento.
Public Nom_Bairro As String         'Bairro.
Public Nom_Cid As String      'Cidade.
Public Sig_UF As String     'UF.
Public Num_CEP As Long        'CEP.
Public Num_Fone As Long         'Telefone.
Public NUM_PI As Double     'PI.

'Ficha 1000307
Public Cod_Reg_Cons As Long             'Código do registro no conselho
Public Num_Reg_Cons As String           'Número do registro

' ======================================================================================================================================================
Private ObjStcPai As stcA46V702B

Private mClsM46V999 As clsM46V999

' ======================================================================================================================================================

Public Property Get Class_Name() As String
    Class_Name = "A46V707B"
End Property

Public Property Get Class_ID() As String
    Class_ID = CStr(nosso_numero)
End Property

Public Property Get gCEP_Formatado() As String
    Dim lAux As String
    If Num_CEP <> 0 Then
        lAux = Format(Num_CEP, "00000000")
        gCEP_Formatado = Mid$(lAux, 1, 5) & "-" & Mid$(lAux, 6, 3)
    Else
        gCEP_Formatado = " "
    End If

End Property

Public Property Set ObjPai(pobjstc As stcA46V702B)
    Set ObjStcPai = pobjstc
End Property

Public Property Get Cod_Unicidade() As Long
    If Not ObjStcPai Is Nothing Then
        Cod_Unicidade = ObjStcPai.Cod_Unicidade
    End If
End Property

Public Property Get gDesc_Reg() As String
    Dim BDP00Tarifa As ADODB.Connection
    Dim rstTbm_Reg As clsYasRecordSet
    Dim lMensagem As String
    Dim lSelect As String

    gDesc_Reg = " "

    If Cod_Reg_Cons = -1 Then
        GoTo SAIDA
    End If

    '1. Abre arquivo.
    If M46V999.gfAbrBasDados("P00MULTI2", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, BDP00Tarifa, lMensagem) = False Then
        Call M46V999.gpGraLog(1, lMensagem)
        GoTo SAIDA
    End If

    '3. Obtém descrição do registro Tbm_Reg_Profissões
    lSelect = "Select Dsc_Reg from Tbm_Reg_Profissoes where Cod_Reg = " & Cod_Reg_Cons

    If M46V999.gfObtRegistro(BDP00Tarifa, lSelect, rstTbm_Reg, lMensagem) = False Then
        Call M46V999.gpGraLog(1, lMensagem)
        GoTo SAIDA
    End If

    If Not rstTbm_Reg.EOF Then
        gDesc_Reg = rstTbm_Reg("Dsc_Reg")
    End If

SAIDA:
    GoSub FechaObj
    Exit Property
FechaObj:
    Set BDP00Tarifa = Nothing
    Set rstTbm_Reg = Nothing
    Return
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
