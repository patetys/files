VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsM99V001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ======================================================================================================================================================
' Class Name: clsM99V001
' Sistema: SyasMult
' Analista: Marcio Oliveira
' Objetivo: Substitui o antigo módulo bas Adaptado para ser utilizado como classe gerenciavel.
' ======================================================================================================================================================
Private Declare Function GetUserName Lib "advapi32.dll" Alias _
                                     "GetUserNameA" (ByVal lpBuffer As String, nsize As Long) As Long

Private Declare Function GetComputerName Lib "kernel32" Alias _
                                         "GetComputerNameA" (ByVal lpBuffer As String, nsize As Long) As Long

Public mOperacao As Integer
Public Cod_Usuario As String
Public Login_Usuario As String
Public Nom_Usuario As String
Public Cod_user_mainframe As String
Public Num_Usuario As String
Public Perfil As String
Public Cod_Grupo_Usuario As String

Public Cons_SEOPE As Boolean
Public Cons_SEDES As Boolean
Public Cons_SESAN As Boolean
Public Cons_SERET As Boolean
Public Cons_PE As Boolean
Public Cons_ADM As Boolean
Public Cons_AplicADM As Boolean
Public Cons_AplicUSU As Boolean
Public UserID As String    'Codigo do DESIS

Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "M99V001"
End Property

Public Function GetLoginComputerName() As String
    Dim cnt As Long
    Dim dl As Long
    Dim S As String

    On Error GoTo 0
    cnt = 200
    S = String(cnt, 0)
    dl = GetComputerName(S, cnt)
    GetLoginComputerName = Left(S, InStr(S, Chr(0)) - 1)
End Function
Public Function GetLoginUserName() As String
    Dim lretorno As Long      'Retorno.
    Dim lUsuario As String    'Usuário.

    lUsuario = Space(64)
    lretorno = GetUserName(lUsuario, 64)
    GetLoginUserName = Left(Trim(lUsuario), Len(Trim(lUsuario)) - 1)
End Function
Public Function Ler_Usuario(DB_base) As Boolean
'Função utilizada para verificar usuário cadastrado no Perfil

    Dim lSelect As String   'Comando SELECT.
    Dim lUsuario As String    'Usuário.
    Dim rsDEPT As DAO.Recordset    'Registro da tabela YTAB_DEPT.
    Dim rsUSER As DAO.Recordset    'Registro da tabela YTAB_USER.
    Dim rsUSER_COMPL As DAO.Recordset    'Registro da tabela YTAB_USER_COMPL.

'=================================================
' Captura: login do usuário       (login_usuario)
'        : código do usuário      (usuario)
'        : código da unidade      (unidade)
'        : nome da unidade        (nome_unidade)
'        : código do departamento (cod_dept)
'=================================================

'1. Prepara controles.
    On Error GoTo ERRO_Ler_Usuario
    Close

    '2. Obtém login do usuário no WINDOWS NT.
    lUsuario = GetLoginUserName()
    'lUsuario = InputBox("Username")
    mOperacao = False
    If UCase(lUsuario) = "OPERACAO" Then
        mOperacao = True
        lUsuario = "juliocr"
    End If
    Cod_Usuario = lUsuario
    If Trim(lUsuario) = "" Then
        Ler_Usuario = False
        On Error GoTo 0
        Exit Function
    End If

    '3. Obtém dados da tabela YTAB_USER.
    Login_Usuario = lUsuario
    Close #1
    lSelect = "SELECT * FROM YTAB_USER WITH (NOLOCK) WHERE COD_LOGIN_USER = '" & UCase(Login_Usuario) & "'"
    Set rsUSER = DB_base.OpenRecordset(lSelect, dbOpenSnapshot)
    If rsUSER.RecordCount = 0 Then
        'MsgBox "Usuário " & Login_Usuario & " não cadastrado!", vbExclamation
        Call M46V999.gpGraLog(1, "Usuário " & Login_Usuario & " não cadastrado!")
        rsUSER.Close
        Set rsUSER = Nothing
        Ler_Usuario = False
        On Error GoTo 0
        Exit Function
    End If
    Cod_Usuario = Trim(rsUSER!Cod_User)
    Nom_Usuario = Trim(rsUSER!cod_login_user)
    Cod_user_mainframe = Trim(rsUSER!Cod_user_mainframe)
    'Perfil = rsUSER!COD_IDEN_PERFIL
    Perfil = rsUSER!COD_PERFIL

    '4. Obtém dados da tabela YTAB_USER_COMPL.
    '-------------------------------------------------------------------------------------------------
    ' Essas linhas foram alteradas em 24/07/2001 em função da conversão do leiaute da tabela YTAB_USER
    ' para torná-lo igual ao da tabela YTAB_USER do mainframe.
    '-------------------------------------------------------------------------------------------------
    lSelect = "SELECT DSC_EMAIL FROM YTAB_USER_COMPL WHERE COD_USER = " & Cod_Usuario
    Set rsUSER_COMPL = DB_base.OpenRecordset(lSelect, dbOpenSnapshot)
    If rsUSER_COMPL.RecordCount = 0 Then
        Num_Usuario = Left(rsUSER!DSC_EMAIL, 3)
    Else
        Num_Usuario = Left(rsUSER_COMPL!DSC_EMAIL, 3)
    End If
    rsUSER_COMPL.Close
    Set rsUSER_COMPL = Nothing
    '-------------------------------------------------------------------------------------------------
    rsUSER.Close
    Set rsUSER = Nothing

    '5. Obtém dados da tabela YTAB_DEPT.
    lSelect = "SELECT COD_GRUPO FROM YTAB_DEPT WHERE NUM_DEPT = " & Mid(Cod_Usuario, 2, 4)
    Set rsDEPT = DB_base.OpenRecordset(lSelect, dbOpenSnapshot)
    If rsDEPT.RecordCount <> 0 Then
        Cod_Grupo_Usuario = Trim(CStr(rsDEPT!cod_grupo))
    End If
    rsDEPT.Close
    Set rsDEPT = Nothing

    '6. Encerra processo.
    On Error GoTo 0
    Ler_Usuario = True
    Exit Function

ERRO_Ler_Usuario:
    'MsgBox "Erro ao obter dados do usuário: " & Err & " - " & Error & ". Consulte o DESIS."
    Call M46V999.gpGraLog(1, "Erro ao obter dados do usuário: " & Err & " - " & ERROR & ". Consulte o DESIS.")
    On Error GoTo 0
End Function


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
