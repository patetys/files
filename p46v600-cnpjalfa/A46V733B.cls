VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V733B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================

' Class Name: stcA46V733B --> Tabela Tab_Ped_ProAgenc
' Sistema: Multi
' Simony Sales

' ======================================================================================================================================================


' ======================================================================================================================================================
Public nosso_numero As String   'Nosso número.
Public NUM_SEQ As Long
Public NUM_PI As Double    'Número do PI.
Public Num_Apol As Double
Public Num_Compl As Long
Public Tip_Canal As Integer
Public Cod_Canal As Long
Public Cod_Corretor As Long
Public Perc_Canal As Double
Public Qtd_Parc_Pag As Integer
Public Qtd_Parc_Apol As Integer
Public Cod_Form_Pag_Ag As Integer
Public Nom_Canal As String
Public Nom_Corr_Canal As String

'CREATE TABLE [dbo].[TAB_CANAL_EMISS](
'    [NOSSO_NUMERO] [varchar](20) NOT NULL,
'    [NUM_SEQ] [numeric](3, 0) NOT NULL,
'    [NUM_PI] [numeric](10, 0) NOT NULL,
'    [NUM_APOL] [numeric](10, 0) NOT NULL,
'    [NUM_COMPL] [numeric](6, 0) NOT NULL,
'
'    [TIP_CANAL] [numeric](2, 0) NOT NULL,
'    [COD_CANAL] [numeric](7, 0) NOT NULL,
'    [COD_CORRETOR] [numeric](7, 0) NOT NULL,
'   [PERC_CANAL] [numeric](5, 3) NOT NULL,
'
'    [QTD_PARC_PAG] [numeric](2, 0) NOT NULL,
'    [QTD_PARC_APOL] [numeric](2, 0) NOT NULL,
'    [COD_FORM_PAG_AG] [numeric](3, 0) NOT NULL,
'    [COD_UNIFICACAO] [numeric](10, 0) NOT NULL,




'======================================================================================================================================================
Private ObjStcPai As stcA46V702B
' ======================================================================================================================================================
Private mNomeCanal As String
Private mCodCanal As Integer

Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "A46V733B"
End Property

Public Property Get Class_ID() As String
    Class_ID = CStr(nosso_numero) & _
               CStr(Format(NUM_SEQ, "0000000"))
End Property

Public Property Get gCopiaProlabore(pNumItem As Long) As stcA46V733B
    Dim lobjstcProlab As stcA46V733B
    Set lobjstcProlab = New stcA46V733B
    Set lobjstcProlab.M46V999 = M46V999

    With lobjstcProlab
        .nosso_numero = nosso_numero
        .NUM_SEQ = NUM_SEQ
        .NUM_PI = NUM_PI
        .Num_Apol = Num_Apol
        .Num_Compl = Num_Compl
        .Tip_Canal = Tip_Canal
        .Cod_Canal = Cod_Canal
        .Cod_Corretor = Cod_Corretor
        .Perc_Canal = Perc_Canal
        .Qtd_Parc_Pag = Qtd_Parc_Pag
        .Qtd_Parc_Apol = Qtd_Parc_Apol
        .Cod_Form_Pag_Ag = Cod_Form_Pag_Ag
    End With

    Set gCopiaProlabore = lobjstcProlab
End Property


Public Property Get gNom_Canal() As String
    Dim lSaida As Variant
    Dim lMensagem As String
    If mCodCanal <> Cod_Canal Or mNomeCanal = "" Then
        If M46V999.gfObtCorr_Pro_Agenc(Me.Tip_Canal, 1, Me.Cod_Canal, lSaida, lMensagem) = True Then
            mNomeCanal = lSaida(1, 3)
            'mNumSusep = lSaida(1, 4)
        Else    'Quando não localizar o corretor ou estiver inavivo.
            mCodCanal = 0
            Cod_Canal = 0
        End If
        mCodCanal = Cod_Canal
    End If
    gNom_Canal = mNomeCanal
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
