VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA28V720B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================
' Class Name: stcA28V720B
' Data/Hora: 12/09/2007 - 11:14:18
' Analista: tjgomes
' ======================================================================================================================================================

' Properties...
' ======================================================================================================================================================
Public NUM_BOLETO As Long    'PK
Public nosso_numero As String
Public COD_UNID As Long
Public Cod_Ramo As Long
Public COD_BANCO As Long
Public Nom_Segurado As String
Public CPF_CNPJ_Segurado As String
Public VAL_PR As Single
Public DAT_VENC As Long
Public Cod_Corr As Long
Public NUM_LINHA_DIGIT As String
Public Cod_User As String
Public DAT_DIGIT As Long
Public HOR_DIGIT As Long
Public STATUS As Long
'*************************************'
Public DigBoleto As Integer
Public Vigencia As String

' ======================================================================================================================================================
Private mCod_Corretor As Long
Private mNome_Corretor As String

Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "A28V720B"
End Property

Public Property Get Class_ID() As String
    Class_ID = CStr(NUM_BOLETO)
End Property

Public Property Get gDigitoBloq() As String
    Dim clsA99V001 As clsA99V001    'DLL: calcula dígitos.
    Dim lDigito As String
    Dim lMensagem As String
    Dim lAuxiliar As String
    Set clsA99V001 = New clsA99V001     'DLL: calcula dígitos.
    Set clsA99V001.M46V999 = M46V999

    lAuxiliar = Format(NUM_BOLETO & "0", "000000000")
    If clsA99V001.DIG_BLOQUETO(lAuxiliar, lDigito, lMensagem) = False Then
        Call M46V999.gpGraLog(1, lMensagem)
    End If
    gDigitoBloq = lDigito
    Set clsA99V001 = Nothing
End Property

Public Property Get gNom_Corretor() As String
    Dim lMensagem As String

    If Me.Cod_Corr <> mCod_Corretor Or mNome_Corretor = "" Then
        mCod_Corretor = Me.Cod_Corr
        If M46V999.gfNomCorretor(mCod_Corretor, mNome_Corretor, lMensagem) = False Then
            mNome_Corretor = ""
            Call M46V999.gpGraLog(1, lMensagem)
            Exit Property
        End If

    End If
    gNom_Corretor = mNome_Corretor
End Property

Public Property Get gLocalImagensBarra() As String
    gLocalImagensBarra = App.Path
End Property
Public Property Get gLocalImagensBanco() As String
    Select Case Me.COD_BANCO
    Case 237
        gLocalImagensBanco = M46V999.gAppPath & "STBS57_237.BMP"
    Case 341
        gLocalImagensBanco = M46V999.gAppPath & "STBS55_341.BMP"
    Case 409
        gLocalImagensBanco = M46V999.gAppPath & "STBS75_409.BMP"
    Case 347
        gLocalImagensBanco = M46V999.gAppPath & "STBS59_347.BMP"
    Case 356
        gLocalImagensBanco = M46V999.gAppPath & "STBS75_356.BMP"
    Case 399
        gLocalImagensBanco = M46V999.gAppPath & "STBS75_399.BMP"
    Case 33
        gLocalImagensBanco = M46V999.gAppPath & "STBS75_033.BMP"
    End Select
End Property

Public Property Get gMsgPagto() As String
    gMsgPagto = Me.COD_UNID & " " & Me.Cod_Corr & " PARC 01/01"
End Property

Public Property Get gobjBoletoCompl() As stcA99V009B
    Dim lobjstcBoleto As stcA99V009B
    Dim lobjBoleto As clsA99V009A
    Dim lConexao As ADODB.Connection
    Dim lMensagem As String
    Set lobjstcBoleto = New stcA99V009B
    Set lobjstcBoleto.M46V999 = M46V999

    With lobjstcBoleto
        .Corretor = Me.gNom_Corretor
        .LocalImagensBarra = Me.gLocalImagensBarra
        .numeroBanco = Me.COD_BANCO
        Select Case .numeroBanco
        Case 237, 341, 356, 399, 33   '09029 alteração boleto santander
            'OK
        Case Else
            .numeroBanco = 341
        End Select
        .LocalImagensBanco = Me.gLocalImagensBanco
        .NomeSacado = Me.Nom_Segurado
        .CPF_CNPJ_Sacado = Me.CPF_CNPJ_Segurado
        .NossoNumero = Me.NUM_BOLETO
        .DigYasNossoNumero = Me.DigBoleto
        .NumeroDoc = Format(Me.NUM_BOLETO & Me.DigBoleto, "0000000-0")
        .ramo = "310 - Automóvel"
        .valorDocumento = Me.VAL_PR
        .vencimento = Me.DAT_VENC
        .Vigencia = Me.Vigencia
        .MensagemPgto = Me.gMsgPagto
        .CarregarDados = True
        If Not M46V999.gfAbrBasDados("P0045300", M46V999.gTipP0045300, M46V999.gEndP0045300, lConexao, lMensagem) Then
            Call M46V999.gpGraLog(1, lMensagem)
            Exit Sub
        End If
        Set .ConexaoBD = lConexao
        .NomeTabela = "P0045300"
    End With

    Set lobjBoleto = New clsA99V009A
    If Not lobjBoleto.mfMontarBoleto(lobjstcBoleto) Then
        GoTo SAIDA
    End If
    Set gobjBoletoCompl = lobjstcBoleto

SAIDA:
    GoSub FechaObj
    Exit Property

FechaObj:
    Set lobjstcBoleto = Nothing
    Set lobjBoleto = Nothing
    Set lConexao = Nothing
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
