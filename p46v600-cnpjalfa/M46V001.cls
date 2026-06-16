VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsM46V001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ======================================================================================================================================================
' Class Name: clsM46V001
' Sistema: SyasMult
' Analista: Marcio Oliveira
' Objetivo: Substitui o antigo m�dulo bas Adaptado para ser utilizado como classe gerenciavel.
' ======================================================================================================================================================

Public gerro As Integer
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nsize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal filename$)

Private mClsM46V111 As clsM46V111

Public Property Get Class_Name() As String
    Class_Name = "M46V001"
End Property

Public Function GravarINI(Sessao As String, NomeChave As String, Valor As String, NomedoArquivo As String)
    Dim rc As Long
    rc = WritePrivateProfileString(Sessao, NomeChave, Valor, NomedoArquivo)
End Function

Public Function LerIni(Secao As String, NomeChave As String, NomedoArquivo As String) As String
'NomedoArquivo=nome do arquivo ini
'Secao=O que esta entre []
'NomeChave=nome do que se encontra antes do sinal de igual

    Dim RetCode As Long
    Dim RetString As String

    LerIni = ""
    RetString = String$(255, 0)
    RetCode = GetPrivateProfileString(Secao, NomeChave, "", RetString, Len(RetString), NomedoArquivo)
    LerIni = Left(RetString, RetCode)

End Function

Public Function mfCriaIndex(pMensagem As String) As Boolean
    Dim lSql As String
    Dim lCreate As String

    ' lSql = " Drop Index Tab_Ped_Cobert_Compl_IND03 on Tab_Ped_Cobert_Compl "
    ' If M46V999.gfExeSQL(bdP00MPED, lSql, pMensagem) <> 0 Then
    '     Call M46V999.gpGraLog(0, pMensagem)
    ' End If
    '
    ' lCreate = ""
    ' lCreate = lCreate & "CREATE  UNIQUE  INDEX   Tab_Ped_Cobert_Compl_IND33"
    ' lCreate = lCreate & "       On Tab_Ped_Cobert_Compl  (NOSSO_NUMERO, NUM_ITEM, COD_COBERT);"
    ' If M46V999.gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
    '     Call M46V999.gpGraLog(1, pMensagem)
    ' End If
    '
    ' lSql = " Drop Index Tab_Ped_Fat_Risco_IND03 on Tab_Ped_Fat_Risco "
    ' If M46V999.gfExeSQL(bdP00MPED, lSql, pMensagem) <> 0 Then
    '     Call M46V999.gpGraLog(0, pMensagem)
    ' End If
    '
    ' lCreate = ""
    ' lCreate = lCreate & "CREATE  UNIQUE  INDEX Tab_Ped_Fat_Risco_IND03"
    ' lCreate = lCreate & "       On Tab_Ped_Fat_Risco (NOSSO_NUMERO, NUM_ITEM, COD_FATOR);"
    ' If M46V999.gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
    '     Call M46V999.gpGraLog(1, pMensagem)
    ' End If
    '
    ' lSql = " Drop Index Tab_Ped_Espec2_IND03 on Tab_Ped_Espec2 "
    ' If M46V999.gfExeSQL(bdP00MPED, lSql, pMensagem) <> 0 Then
    '     Call M46V999.gpGraLog(0, pMensagem)
    ' End If
    '
    '
    '  lCreate = ""
    '  lCreate = lCreate & "CREATE  UNIQUE  INDEX Tab_Ped_Espec2_IND03"
    '  lCreate = lCreate & "       On Tab_Ped_Espec2 (NOSSO_NUMERO);"
    '  If M46V999.gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
    '      Call M46V999.gpGraLog(1, pMensagem)
    '  End If
    '
    '
    ' lSql = " Drop Index Tab_Ped_Espec_IND03 on Tab_Ped_Espec "
    ' If M46V999.gfExeSQL(bdP00MPED, lSql, pMensagem) <> 0 Then
    '     Call M46V999.gpGraLog(0, pMensagem)
    ' End If
    '
    ' lCreate = ""
    ' lCreate = lCreate & "CREATE  UNIQUE  INDEX Tab_Ped_Espec_IND03"
    ' lCreate = lCreate & "       On Tab_Ped_Espec (NOSSO_NUMERO, NUM_PAG);"
    'If M46V999.gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
    '     Call M46V999.gpGraLog(1, pMensagem)
    'End If'
    '


End Function

Public Function mfCriaTabelas(pMensagem As String) As Boolean
    On Error GoTo CriaTabelas_Err

    Dim lCreate As String
    Dim lAlter As String
    Dim bdP00MPED As ADODB.Connection    'Base de dados: P00MPED.
    Dim lrsTabela As clsYasRecordSet    'Registro da tabela.
    Dim lSql As String
    Dim bdP00MULTI As ADODB.Connection    'Base de dados: P00MULTI.


    If M46V999.gTipUsuario = e_Tipo_Usu_Yasuda And M46V999.gBaseLocal = False Then
        mfCriaTabelas = True
        Exit Function
    End If
    Screen.MousePointer = vbHourglass

    mfCriaTabelas = False
    '1. Abre base de dados.
    If M46V999.gfAbrBasDados("P00MPED", M46V999.gTipP00MPED, M46V999.gEndP00MPED, bdP00MPED, pMensagem) = False Then
        Call M46V999.gpGraLog(2, pMensagem)
        Exit Function
        'End
    End If

    If M46V999.gBaseLocal = True Then

        lSql = "Select top 1 * from Tab_Ctrl_Emis_Compl WITH (NOLOCK) "
        pMensagem = ""
        If Not M46V999.gfObtRegistro(bdP00MPED, lSql, lrsTabela, pMensagem) Then

            lCreate = ""
            lCreate = lCreate & "    CREATE TABLE Tab_Ctrl_Emis_Compl ("
            lCreate = lCreate & "          nosso_numero varchar(20) PRIMARY KEY NOT NULL,"
            lCreate = lCreate & "          num_pi INT NOT NULL,"
            lCreate = lCreate & "          cod_ramo INT NOT NULL,"
            lCreate = lCreate & "          num_prop INT NOT NULL,"
            lCreate = lCreate & "          num_apol INT NOT NULL,"
            lCreate = lCreate & "          num_endo INT NOT NULL,"
            lCreate = lCreate & "          Parcelas_Impressao_Avista varchar(50) NULL,"
            lCreate = lCreate & "          Parcelas_Impressao_30dias varchar(50) NULL,"
            lCreate = lCreate & "          Parcelas_Impressao_AvistaDebito varchar(50) NULL,"
            lCreate = lCreate & "          Parcelas_Impressao_30diasDebito varchar(50) NULL,"
            lCreate = lCreate & "          Parcelas_Impressao_Cartao varchar(50) NULL,"
            lCreate = lCreate & "          Email_Contato varchar(150) NULL,"
            lCreate = lCreate & "          Ind_Impr_Apol SHORT  NOT NULL,"
            lCreate = lCreate & "          Ind_Prolabore SHORT NOT NULL,"
            lCreate = lCreate & "          Ind_Agenciamento INT NOT NULL,"
            lCreate = lCreate & "          dat_digit INT NOT NULL,"
            lCreate = lCreate & "          hor_digit INT NOT NULL,"
            lCreate = lCreate & "          Cod_Bco_Boleto int NULL,"
            lCreate = lCreate & "          Cod_Origem INT NULL,"
            lCreate = lCreate & "          Num_Apol_Mar double NULL);"

            If M46V999.gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
                Call M46V999.gpGraLog(1, pMensagem)
            End If

        End If

    End If

    lSql = "Select top 1 * from Tab_Ped_Cobert_Compl WITH (NOLOCK) "
    pMensagem = ""
    If Not M46V999.gfObtRegistro(bdP00MPED, lSql, lrsTabela, pMensagem) Then
        lCreate = ""
        lCreate = lCreate & "CREATE TABLE  Tab_Ped_Cobert_Compl ("
        lCreate = lCreate & "     NOSSO_NUMERO   char  (20) Null,"
        lCreate = lCreate & "     COD_ORG   int  NOT NULL ,"
        lCreate = lCreate & "     COD_RAMO   int NOT NULL ,"
        lCreate = lCreate & "     NUM_PED   double NOT NULL ,"
        lCreate = lCreate & "     NUM_APOL   double  NOT NULL ,"
        lCreate = lCreate & "     NUM_ENDO   double  NOT NULL ,"
        lCreate = lCreate & "     NUM_ITEM   double  NOT NULL ,"
        lCreate = lCreate & "     TAB_COBERT int NOT NULL ,"
        lCreate = lCreate & "     COD_COBERT int NOT NULL ,"
        lCreate = lCreate & "     Num_Meses_Pi int NULL ,"
        lCreate = lCreate & "     Formula   char  (100) Null,"
        lCreate = lCreate & "     Cod_Classe int NULL ,"
        lCreate = lCreate & "     Val_Fator_Franq   double   NULL ,"
        lCreate = lCreate & "     Alterar_Cobert   char (10)  NULL ,"
        lCreate = lCreate & "     Val_Risco_Cob   double   NULL ,"
        lCreate = lCreate & "     NUM_PI   double  NULL, "
        lCreate = lCreate & "     AjusteEndosso char(5) null,"
        lCreate = lCreate & "     Alteroufranquia char(3) null,"
        lCreate = lCreate & "     TextoFranquia char(200) null, "
        lCreate = lCreate & "     Val_Pr_Orig double  NULL, "
        lCreate = lCreate & "     PremioInformado char(3)  NULL"
        'ProjetoFormularios
        lCreate = lCreate & "     ,Compl_Desc_Cobertura varchar(50)  NULL"
        'Regra de Batente
        lCreate = lCreate & "     ,Val_Pr_Batente decimal(15,2) NULL"
        'Multi Franquia
        lCreate = lCreate & "     ,Per_Desc_Franq decimal(4,2) NULL"
        lCreate = lCreate & "     )  ;"

        If M46V999.gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If

        lSql = " Drop Index Tab_Ped_Cobert_Compl_IND03 on Tab_Ped_Cobert_Compl "
        If M46V999.gfExeSQL(bdP00MPED, lSql, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(0, pMensagem)
        End If


        lCreate = ""
        lCreate = lCreate & "CREATE  UNIQUE  INDEX   Tab_Ped_Cobert_Compl_IND03"
        lCreate = lCreate & "       On Tab_Ped_Cobert_Compl  (NUM_PI, NUM_ITEM, COD_COBERT);"
        If M46V999.gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If

        lCreate = ""
        lCreate = lCreate & "CREATE    INDEX     Tab_Ped_Cobert_Compl_IND01"
        lCreate = lCreate & "       On Tab_Ped_Cobert_Compl (COD_ORG,COD_RAMO,NUM_PED) ;"
        If M46V999.gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If

        lCreate = ""
        lCreate = lCreate & "CREATE    INDEX   Tab_Ped_Cobert_Compl_IND02"
        lCreate = lCreate & "       On Tab_Ped_Cobert_Compl  (NUM_APOL,NUM_ENDO,NUM_ITEM);"
        If M46V999.gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If
    End If

    lSql = "Select top 1 AjusteEndosso from Tab_Ped_Cobert_Compl WITH (NOLOCK) "
    pMensagem = ""
    If Not M46V999.gfObtRegistro(bdP00MPED, lSql, lrsTabela, pMensagem) Then
        lAlter = " Alter table Tab_Ped_Cobert_Compl add COLUMN AjusteEndosso varchar(05); "
        If M46V999.gfExeSQL(bdP00MPED, lAlter, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If

        lSql = " Update Tab_Ped_Cobert_Compl "
        lSql = lSql & " set AjusteEndosso = '     ' "
        If M46V999.gfExeSQL(bdP00MPED, lSql, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If
    End If

    lSql = "Select top 1 TextoFranquia from Tab_Ped_Cobert_Compl WITH (NOLOCK) "
    pMensagem = ""
    If Not M46V999.gfObtRegistro(bdP00MPED, lSql, lrsTabela, pMensagem) Then
        lAlter = " Alter table Tab_Ped_Cobert_Compl add COLUMN TextoFranquia varchar(200); "
        If M46V999.gfExeSQL(bdP00MPED, lAlter, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If

        lSql = " Update Tab_Ped_Cobert_Compl "
        lSql = lSql & " set TextoFranquia = '     ' "
        If M46V999.gfExeSQL(bdP00MPED, lSql, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If
    End If


    lSql = "Select top 1 AlterouFranquia from Tab_Ped_Cobert_Compl WITH (NOLOCK) "
    pMensagem = ""
    If Not M46V999.gfObtRegistro(bdP00MPED, lSql, lrsTabela, pMensagem) Then
        lAlter = " Alter table Tab_Ped_Cobert_Compl add COLUMN AlterouFranquia varchar(03); "
        If M46V999.gfExeSQL(bdP00MPED, lAlter, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If

        lSql = " Update Tab_Ped_Cobert_Compl "
        lSql = lSql & " set AlterouFranquia = '   ' "
        If M46V999.gfExeSQL(bdP00MPED, lSql, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If
    End If

    ' 1300639 - Pr�mio Informado - Inicio
    lSql = "Select top 1 Val_PR_Orig from Tab_Ped_Cobert_Compl WITH (NOLOCK) "
    pMensagem = ""
    If Not M46V999.gfObtRegistro(bdP00MPED, lSql, lrsTabela, pMensagem) Then
        lAlter = " Alter table Tab_Ped_Cobert_Compl add COLUMN Val_PR_Orig double; "
        If M46V999.gfExeSQL(bdP00MPED, lAlter, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If

        lAlter = " Alter table Tab_Ped_Cobert_Compl add COLUMN PremioInformado varchar(03) ; "
        If M46V999.gfExeSQL(bdP00MPED, lAlter, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If

        'ProjetoFormularios
        lAlter = " Alter table Tab_Ped_Cobert_Compl add COLUMN Compl_Desc_Cobertura varchar(50) ; "
        If M46V999.gfExeSQL(bdP00MPED, lAlter, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If

        'Regra de Batente
        lAlter = " Alter table Tab_Ped_Cobert_Compl add COLUMN Val_Pr_Batente decimal(15,2) ; "
        If M46V999.gfExeSQL(bdP00MPED, lAlter, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If

        'Multiplo de Franquia
        lAlter = " Alter table Tab_Ped_Cobert_Compl add COLUMN Per_Desc_Franq decimal(4,2) ; "
        If M46V999.gfExeSQL(bdP00MPED, lAlter, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If


        lSql = " Update Tab_Ped_Cobert_Compl "
        lSql = lSql & " set PremioInformado = 0 "
        If M46V999.gfExeSQL(bdP00MPED, lSql, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If
    End If
    ' 1300639 - Pr�mio Informado - Fim



    lSql = "Select top 1 * from Tab_Ped_Fat_Risco WITH (NOLOCK) "
    pMensagem = ""
    If Not M46V999.gfObtRegistro(bdP00MPED, lSql, lrsTabela, pMensagem) Then
        lCreate = ""
        lCreate = lCreate & "CREATE TABLE  Tab_Ped_Fat_Risco ("
        lCreate = lCreate & "     NOSSO_NUMERO   char  (20) Null,"
        lCreate = lCreate & "     COD_ORG   int NOT NULL ,"
        lCreate = lCreate & "     COD_RAMO   int NOT NULL ,"
        lCreate = lCreate & "     NUM_PED   int NOT NULL ,"
        lCreate = lCreate & "     NUM_APOL   double NOT NULL ,"
        lCreate = lCreate & "     NUM_ENDO   long NOT NULL ,"
        lCreate = lCreate & "     NUM_ITEM   long NOT NULL ,"
        lCreate = lCreate & "     COD_FATOR   int NOT NULL ,"
        lCreate = lCreate & "     DSC_COMPL_FATOR   char  (30),"
        lCreate = lCreate & "     NUM_PI   double   NULL )  ;"
        If M46V999.gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If


        lCreate = ""
        lCreate = lCreate & "CREATE  UNIQUE  INDEX Tab_Ped_Fat_Risco_IND03"
        lCreate = lCreate & "       On Tab_Ped_Fat_Risco (NUM_PI, NUM_ITEM, COD_FATOR);"
        If M46V999.gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If

        lCreate = ""
        lCreate = lCreate & "CREATE    INDEX Tab_Ped_Fat_Risco_IND01"
        lCreate = lCreate & "       On Tab_Ped_Fat_Risco (COD_ORG,COD_RAMO,NUM_PED, COD_FATOR) ;"
        If M46V999.gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If

        lCreate = ""
        lCreate = lCreate & "CREATE    INDEX Tab_Ped_Fat_Risco_IND02"
        lCreate = lCreate & "       On Tab_Ped_Fat_Risco (NUM_APOL,NUM_ENDO,NUM_ITEM, COD_FATOR);"
        If M46V999.gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If
    End If

    lSql = "Select top 1 * from Tab_Ped_Espec2 WITH (NOLOCK) "
    pMensagem = ""
    If Not M46V999.gfObtRegistro(bdP00MPED, lSql, lrsTabela, pMensagem) Then
        lCreate = ""
        lCreate = lCreate & "CREATE TABLE  Tab_Ped_Espec2 ("
        lCreate = lCreate & "     NOSSO_NUMERO   char  (20) Null,"
        lCreate = lCreate & "     COD_ORG   int NOT NULL ,"
        lCreate = lCreate & "     COD_RAMO   int NOT NULL ,"
        lCreate = lCreate & "     NUM_PED   int NOT NULL ,"
        lCreate = lCreate & "     NUM_APOL   double NOT NULL ,"
        lCreate = lCreate & "     NUM_ENDO   int NOT NULL ,"
        lCreate = lCreate & "     DSC_OBJETOS   text  Null,"
        lCreate = lCreate & "     DSC_RCOBERTOS   text  Null,"
        lCreate = lCreate & "     DSC_CARACTERISTICA   text  Null,"
        lCreate = lCreate & "     DSC_PERIMETRO   text  Null,"
        lCreate = lCreate & "     DSC_ESPECIFICACAO   text  Null,"
        lCreate = lCreate & "     DSC_COBERTURAS   text  Null,"
        lCreate = lCreate & "     DSC_FRANQUIA   text  Null,"
        lCreate = lCreate & "     NOM_CONTRATANTE   char  (50) Null,"
        lCreate = lCreate & "     DSC_CLAUSULAS   text  Null,"
        lCreate = lCreate & "     DSC_CONDICAO   text  Null,"
        lCreate = lCreate & "     DSC_OUTRASINF   text  Null,"
        lCreate = lCreate & "     DSC_OBSERVACAO   text  Null,"
        lCreate = lCreate & "     DAT_OBRA   int  NULL ,"
        lCreate = lCreate & "     DAT_MANUT   int NULL ,"
        lCreate = lCreate & "     NUM_PI   double NOT NULL , "
        lCreate = lCreate & "     IND_IMPRESSAO int NULL, "
        lCreate = lCreate & "     Dsc_LocalExposicao   text  Null ) ;"

        If M46V999.gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If

        lCreate = ""
        lCreate = lCreate & "CREATE  UNIQUE  INDEX Tab_Ped_Espec2_IND03"
        lCreate = lCreate & "       On Tab_Ped_Espec2 (NUM_PI);"
        If M46V999.gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If

        lCreate = ""
        lCreate = lCreate & "CREATE    INDEX Tab_Ped_Espec2_IND01"
        lCreate = lCreate & "       On Tab_Ped_Espec2 (COD_ORG,COD_RAMO,NUM_PED) ;"
        If M46V999.gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If

        lCreate = ""
        lCreate = lCreate & "CREATE    INDEX Tab_Ped_Espec2_IND02"
        lCreate = lCreate & "       On Tab_Ped_Espec2 (NUM_APOL,NUM_ENDO);"
        If M46V999.gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If
    End If

    lSql = "Select top 1 * from Tab_Ped_Espec WITH (NOLOCK) "
    pMensagem = ""
    If Not M46V999.gfObtRegistro(bdP00MPED, lSql, lrsTabela, pMensagem) Then
        lCreate = ""
        lCreate = lCreate & "CREATE TABLE  Tab_Ped_Espec ("
        lCreate = lCreate & "     Nosso_Numero   char  (20) Null,"
        lCreate = lCreate & "     COD_ESPEC   int   NULL ,"
        lCreate = lCreate & "     COD_ORG int NOT NULL ,"
        lCreate = lCreate & "     COD_RAMO int NOT NULL ,"
        lCreate = lCreate & "     NUM_PED int NOT NULL ,"
        lCreate = lCreate & "     NUM_APOL   double  NOT NULL ,"
        lCreate = lCreate & "     NUM_ENDO int NOT NULL ,"
        lCreate = lCreate & "     NUM_PAG int NULL ,"
        lCreate = lCreate & "     QTD_VIAS int NULL ,"
        lCreate = lCreate & "     ESPEC_LIN1   char  (80) Null,"
        lCreate = lCreate & "     ESPEC_LIN2   char  (80) Null,"
        lCreate = lCreate & "     ESPEC_LIN3   char  (80) Null,"
        lCreate = lCreate & "     ESPEC_LIN4   char  (80) Null,"
        lCreate = lCreate & "     ESPEC_LIN5   char  (80) Null,"
        lCreate = lCreate & "     ESPEC_LIN6   char  (80) Null,"
        lCreate = lCreate & "     ESPEC_LIN7   char  (80) Null,"
        lCreate = lCreate & "     ESPEC_LIN8   char  (80) Null,"
        lCreate = lCreate & "     ESPEC_LIN9   char  (80) Null,"
        lCreate = lCreate & "     ESPEC_LIN10   char  (80) Null,"
        lCreate = lCreate & "     ESPEC_LIN11   char  (80) Null,"
        lCreate = lCreate & "     ESPEC_LIN12   char  (80) Null,"
        lCreate = lCreate & "     ESPEC_LIN13   char  (80) Null,"
        lCreate = lCreate & "     ESPEC_LIN14   char  (80) Null,"
        lCreate = lCreate & "     ESPEC_LIN15   char  (80) Null,"
        lCreate = lCreate & "     ESPEC_LIN16   char  (80) Null,"
        lCreate = lCreate & "     ESPEC_LIN17   char  (80) Null,"
        lCreate = lCreate & "     ESPEC_LIN18   char  (80) Null,"
        lCreate = lCreate & "     ESPEC_LIN19   char  (80) Null,"
        lCreate = lCreate & "     ESPEC_LIN20   char  (80) Null,"
        lCreate = lCreate & "     NUM_PI int NOT NULL ) ;"
        If M46V999.gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If

        lCreate = ""
        lCreate = lCreate & "CREATE  UNIQUE  INDEX Tab_Ped_Espec_IND03"
        lCreate = lCreate & "       On Tab_Ped_Espec (NUM_PI, NUM_PAG);"
        If M46V999.gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If

        lCreate = ""
        lCreate = lCreate & "CREATE    INDEX Tab_Ped_Espec_IND01"
        lCreate = lCreate & "       On Tab_Ped_Espec (COD_ORG,COD_RAMO,NUM_PED, NUM_PAG) ;"
        If M46V999.gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If

        lCreate = ""
        lCreate = lCreate & "CREATE    INDEX Tab_Ped_Espec_IND02"
        lCreate = lCreate & "       On Tab_Ped_Espec (NUM_APOL,NUM_ENDO, NUM_PAG);"
        If M46V999.gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If
    End If

    lSql = "Select top 1 * from Tab_Ped_Dizer WITH (NOLOCK) "
    pMensagem = ""
    If Not M46V999.gfObtRegistro(bdP00MPED, lSql, lrsTabela, pMensagem) Then
        lCreate = ""
        lCreate = lCreate & "CREATE TABLE  Tab_Ped_Dizer  ("
        lCreate = lCreate & "     NOSSO_NUMERO   char  (20)   NOT NULL ,"
        lCreate = lCreate & "     COD_ORG   int NOT NULL ,"
        lCreate = lCreate & "     COD_RAMO   int NOT NULL ,"
        lCreate = lCreate & "     NUM_PED   int NOT NULL ,"
        lCreate = lCreate & "     NUM_APOL   double NOT NULL ,"
        lCreate = lCreate & "     NUM_ENDO   int  NOT NULL ,"
        lCreate = lCreate & "     NUM_SEQ   int NOT NULL ,"
        lCreate = lCreate & "     DSC_DIZER_ENDO   char (80)   NOT NULL )"
        If M46V999.gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If

        lCreate = ""
        lCreate = lCreate & " CREATE  INDEX  TAB_PED_DIZER_IND01  ON  Tab_Ped_Dizer"
        lCreate = lCreate & "( NUM_PED ,  COD_RAMO ,  COD_ORG )"
        If M46V999.gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If

        lCreate = ""
        lCreate = lCreate & " CREATE  INDEX  TAB_PED_DIZER_IND02  ON   Tab_Ped_Dizer"
        lCreate = lCreate & "       ( NUM_APOL ,  NUM_ENDO )"
        If M46V999.gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If
    End If

    lSql = "Select top 1 * from Tab_Ped_Auto WITH (NOLOCK) "
    pMensagem = ""
    If Not M46V999.gfObtRegistro(bdP00MPED, lSql, lrsTabela, pMensagem) Then
        lCreate = ""
        lCreate = lCreate & "CREATE TABLE  Tab_Ped_Auto  ("
        lCreate = lCreate & "     NOSSO_NUMERO    char  (20)   NOT NULL ,"
        lCreate = lCreate & "     COD_ORG   int NOT NULL ,"
        lCreate = lCreate & "     COD_RAMO   int NOT NULL ,"
        lCreate = lCreate & "     NUM_PED   int NOT NULL ,"
        lCreate = lCreate & "     NUM_APOL   double NOT NULL ,"
        lCreate = lCreate & "     NUM_ENDO   int  NOT NULL ,"
        lCreate = lCreate & "     VAL_IS_AUTO double NOT NULL ,"
        lCreate = lCreate & "     VAL_IS_RC double NOT NULL ,"
        lCreate = lCreate & "     VAL_IS_APP double NOT NULL ,"
        lCreate = lCreate & "     VAL_PR_AUTO double NOT NULL ,"
        lCreate = lCreate & "     VAL_PR_RC double NOT NULL ,"
        lCreate = lCreate & "     VAL_PR_APP double NOT NULL ,"
        lCreate = lCreate & "     NUM_PI double NOT NULL )"
        If M46V999.gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If

        lCreate = ""
        lCreate = lCreate & " CREATE  INDEX  TAB_PED_AUTO_IND01  ON Tab_Ped_Auto"
        lCreate = lCreate & "          ( COD_ORG ,  COD_RAMO ,  NUM_PED )"
        If M46V999.gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If

        lCreate = ""
        lCreate = lCreate & " CREATE  INDEX  TAB_PED_AUTO_IND02  ON  Tab_Ped_Auto"
        lCreate = lCreate & "           ( NUM_APOL ,  NUM_ENDO ) "
        If M46V999.gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If
    End If

    lSql = "Select top 1 * from Tbm_Ramo WITH (NOLOCK) "
    pMensagem = ""
    If Not M46V999.gfObtRegistro(bdP00MPED, lSql, lrsTabela, pMensagem) Then
        lCreate = ""
        lCreate = lCreate & "CREATE TABLE  Tbm_Ramo  ("
        lCreate = lCreate & "     COD_RAMO   int  NOT NULL ,"
        lCreate = lCreate & "     DSC_RAMO   char  (30)   NULL ,"
        lCreate = lCreate & "     Tipo_Usuario  char  (1)   NULL )"
        If M46V999.gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If

        lSql = "Insert into Tbm_Ramo values (112, 'Empresarial','C') "
        If M46V999.gfExeSQL(bdP00MPED, lSql, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If
        lSql = "Insert into Tbm_Ramo values (113, 'Residencial','C') "
        If M46V999.gfExeSQL(bdP00MPED, lSql, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If
        lSql = "Insert into Tbm_Ramo values (114, 'Condom�nio','C') "
        If M46V999.gfExeSQL(bdP00MPED, lSql, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If
    End If

    lSql = "Select top 1 Cod_Bem, Cod_Seq from Tab_Ped_Loc_Bem WITH (NOLOCK) "
    pMensagem = ""
    If Not M46V999.gfObtRegistro(bdP00MPED, lSql, lrsTabela, pMensagem) Then
        lSql = " alter table Tab_Ped_Loc_Bem add COLUMN Cod_Bem int, "
        lSql = lSql & " Cod_Seq int"
        If M46V999.gfExeSQL(bdP00MPED, lSql, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If
        lSql = " Update Tab_Ped_Loc_Bem"
        lSql = lSql & " set Cod_Bem = 999 , Cod_Seq = 1  "
        If M46V999.gfExeSQL(bdP00MPED, lSql, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If
        lSql = " Drop Index PrimaryKey on Tab_Ped_Loc_Bem "
        If M46V999.gfExeSQL(bdP00MPED, lSql, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(0, pMensagem)
        End If
        lSql = " Create Unique Index PrimaryKey on Tab_Ped_Loc_Bem "
        lSql = lSql & " (Nosso_Numero, Num_Item,  Cod_Bem, Cod_Seq, Dsc_Bem )"
        If M46V999.gfExeSQL(bdP00MPED, lSql, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If
    End If

    lSql = "Select top 1 * from Tab_Ped_Num_Copia WITH (NOLOCK) "
    pMensagem = ""
    If Not M46V999.gfObtRegistro(bdP00MPED, lSql, lrsTabela, pMensagem) Then
        lCreate = ""
        lCreate = lCreate & "CREATE TABLE Tab_Ped_Num_Copia ("
        lCreate = lCreate & "   NOSSO_NUMERO_NEW  char (20) NOT NULL ,"
        lCreate = lCreate & "   NOSSO_NUMERO_OLD  char (20) NOT NULL ,"
        lCreate = lCreate & "   DAT_COPIA  double NOT NULL )"

        If M46V999.gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If

        lCreate = ""
        lCreate = lCreate & "  CREATE  INDEX Tab_Ped_Num_Copia_IND01  ON "
        lCreate = lCreate & " Tab_Ped_Num_Copia (NOSSO_NUMERO_NEW) "

        If M46V999.gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If
    End If


    '    '2. Abre base de dados 2.
    '    If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, bdP00MULTI, pMensagem) = False Then
    '        Call M46V999.gpGraLog(2, pMensagem)
    '        End
    '    End If
    '
    '    lSql = "Select * from Tbm_Taxa_Classe where "
    '    lSql = lSql & " COD_RAMO = 113 and "
    '    lSql = lSql & " COD_MODALIDADE = 1 and "
    '    lSql = lSql & " COD_Faixa_Reg = 3 and "
    '    lSql = lSql & " COD_CLASSE_LMG = 0 and "
    '    lSql = lSql & " COD_CLASSE = 3 and "
    '    lSql = lSql & " COD_CONSTRUCAO = 2 and "
    '    lSql = lSql & " COD_TARIF = 922 and "
    '    lSql = lSql & " COD_COBERT = 5 and "
    '    lSql = lSql & " Cod_Sequencia  = 0 and "
    ''Ficha 1000305
    '    lSql = lSql & " Tip_Cobert = 0 "
    '    pMensagem = ""
    '    If Not M46V999.gfObtRegistro(bdP00MULTI, lSql, lrsTabela, pMensagem) Then
    '        Call M46V999.gpGraLog(2, pMensagem)
    '        End
    '    End If
    '    If lrsTabela.EOF Then
    '        lSql = "Insert into Tbm_Taxa_Classe values ("
    ''Ficha 1000305
    '        lSql = lSql & "113,1,3,0,3,2,5,5.00000,5.00000,5.00000,922,0,0 ) "
    '        If M46V999.gfExeSQL(bdP00MULTI, lSql, pMensagem) <> 0 Then
    '            Call M46V999.gpGraLog(1, pMensagem)
    '        End If
    '    End If

    ' INICIO - Tarifa 1056 e 1057
    If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, bdP00MULTI, pMensagem) = False Then
        Call M46V999.gpGraLog(2, pMensagem)
        GoTo CriaTabelas_Err
    End If

    lSql = "Select top 1 * from Tbm_Per_Ind_Padrao WITH (NOLOCK) "
    pMensagem = ""
    If Not M46V999.gfObtRegistro(bdP00MULTI, lSql, lrsTabela, pMensagem) Then
        lCreate = ""
        lCreate = lCreate & "CREATE TABLE Tbm_Per_Ind_Padrao ("
        lCreate = lCreate & "   COD_RAMO        INT NOT NULL ,"
        lCreate = lCreate & "   COD_COBERT      INT NOT NULL ,"
        lCreate = lCreate & "   COD_TARIF       INT NOT NULL ,"
        lCreate = lCreate & "   COD_SEQUENCIA   INT NOT NULL ,"
        lCreate = lCreate & "   Per_Indenit_Def INT NOT NULL )"

        If M46V999.gfExeSQL(bdP00MULTI, lCreate, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If
    End If

    lSql = "Select top 1 * from Tbm_VR_LMG_Agrav WITH (NOLOCK) "
    pMensagem = ""
    If Not M46V999.gfObtRegistro(bdP00MULTI, lSql, lrsTabela, pMensagem) Then
        lCreate = ""
        lCreate = lCreate & "CREATE TABLE Tbm_VR_LMG_Agrav ("
        lCreate = lCreate & "   COD_RAMO        INT NOT NULL ,"
        lCreate = lCreate & "   COD_COBERT      INT NOT NULL ,"
        lCreate = lCreate & "   COD_TARIF       INT NOT NULL ,"
        lCreate = lCreate & "   PROP_LMG_VR     INT NOT NULL ,"
        lCreate = lCreate & "   Percent_Agravo  DOUBLE NOT NULL )"

        If M46V999.gfExeSQL(bdP00MULTI, lCreate, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If
    End If
    ' FIM - Tarifa 1056 e 1057


    lSql = "Select top 1 * from Tab_Ped_ProAgenc WITH (NOLOCK) "
    pMensagem = ""
    If Not M46V999.gfObtRegistro(bdP00MPED, lSql, lrsTabela, pMensagem) Then
        lCreate = ""
        lCreate = lCreate & "CREATE TABLE Tab_Ped_ProAgenc ("
        lCreate = lCreate & "   NOSSO_NUMERO   char  (20)  NOT NULL,"
        lCreate = lCreate & "   COD_RAMO   int NOT NULL , "
        lCreate = lCreate & "   COD_CORR  int NOT NULL , "
        lCreate = lCreate & "   PER_CORR  double  NOT NULL , "
        lCreate = lCreate & "   NUM_ORDEM_CORR int NOT NULL, "
        lCreate = lCreate & "   NUM_PI double NOT NULL, "
        lCreate = lCreate & "   TIP_CORR char(2) NOT NULL, "
        lCreate = lCreate & "   TIP_PAGTO char(2) NOT NULL) "


        If M46V999.gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If

        lCreate = ""
        lCreate = lCreate & "CREATE  UNIQUE  INDEX   Tab_Ped_ProAgenc_IND01"
        lCreate = lCreate & "       On Tab_Ped_ProAgenc  (NUM_PI, COD_RAMO, COD_CORR);"
        If M46V999.gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If

        lCreate = ""
        lCreate = lCreate & "CREATE    INDEX     Tab_Ped_ProAgenc_IND02"
        lCreate = lCreate & "       On Tab_Ped_ProAgenc (NOSSO_NUMERO,COD_CORR) ;"
        If M46V999.gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If

    End If

    lSql = "Select top 1 * from Tab_Ped_Loc_Compl WITH (NOLOCK) "
    pMensagem = ""
    If Not M46V999.gfObtRegistro(bdP00MPED, lSql, lrsTabela, pMensagem) Then
        lCreate = ""
        lCreate = lCreate & "CREATE TABLE Tab_Ped_Loc_Compl ("
        lCreate = lCreate & "   NOSSO_NUMERO    char    (20)    NOT NULL,"
        lCreate = lCreate & "   COD_ORG         int             NOT NULL,"
        lCreate = lCreate & "   COD_RAMO        int             NOT NULL,"
        lCreate = lCreate & "   NUM_PED         double          NOT NULL,"
        lCreate = lCreate & "   NUM_APOL        double          NOT NULL,"
        lCreate = lCreate & "   NUM_ENDO        double          NOT NULL,"
        lCreate = lCreate & "   NUM_ITEM        double          NOT NULL,"
        lCreate = lCreate & "   Ind_VerbaSep    double          NOT NULL,"
        lCreate = lCreate & "   Val_Vr_Lc       double          NOT NULL,"
        lCreate = lCreate & "   Val_LMG_Lc      double          NOT NULL,"
        lCreate = lCreate & "   Per_IS_Vr_Lc    double          NOT NULL,"
        lCreate = lCreate & "   Coef_IS_Vr_Lc   double          NOT NULL)"

        If M46V999.gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If
    End If

    lSql = "Select top 1 * from Tbm_Plano_Vida WITH (NOLOCK) "
    pMensagem = ""
    If Not M46V999.gfObtRegistro(bdP00MULTI, lSql, lrsTabela, pMensagem) Then
        lCreate = ""
        lCreate = lCreate & "CREATE TABLE Tbm_Plano_Vida ("
        lCreate = lCreate & "   COD_RAMO        int             NOT NULL,"
        lCreate = lCreate & "   COD_MODALIDADE  int             NOT NULL,"
        lCreate = lCreate & "   COD_PLANO_VIDA  int             NOT NULL,"
        lCreate = lCreate & "   DSC_PLANO_VIDA  char    (20)    NOT NULL,"
        lCreate = lCreate & "   TIP_PLANO_VIDA  int             NOT NULL,"
        lCreate = lCreate & "   COD_COBERT      int             NOT NULL,"
        lCreate = lCreate & "   COD_TARIF       INT             NOT NULL,"
        lCreate = lCreate & "   COD_SEQUENCIA   INT             NOT NULL,"
        lCreate = lCreate & "   VAL_LMG_PADRAO  double          NOT NULL)"

        If M46V999.gfExeSQL(bdP00MULTI, lCreate, pMensagem) <> 0 Then
            Call M46V999.gpGraLog(1, pMensagem)
        End If
    End If

    Call M46V999.gpFecha2(bdP00MULTI)
    Call M46V999.gpFecha2(bdP00MPED)
    Set bdP00MPED = Nothing
    Set lrsTabela = Nothing
    Set bdP00MULTI = Nothing


    mfCriaTabelas = True
CriaTabelas_Err:
End Function


Private Sub mpCarregaCorretor(ByRef pCod_Corretor As Long)
'Fun��o: Busca o corretor na arquivo P2842400.EXE

'Par�metro de sa�da.....pCod_Corretor......C�digo do corretor.

    Dim lArquivo As String    'Endere�o do arquivo.
    Dim lNumArquivo As Integer      'Utilizado com FreeFile.
    Dim lRegistro As String     'Registro.

'1. Inicio a vari�vel.
    pCod_Corretor = 0

    '2. Busco o dados no arquivo P2842400.EXE
    pCod_Corretor = 0
    lArquivo = M46V999.gAppPath & "P2842400.EXE"
    If M46V999.gfPreenchido(Dir(lArquivo)) = False Then
        lArquivo = M46V999.gAppPath & "P2842401.EXE"
        If M46V999.gfPreenchido(Dir(lArquivo)) = False Then
            Exit Sub
        End If
    End If
    lNumArquivo = FreeFile
    On Error Resume Next
    Open lArquivo For Input As #lNumArquivo
    If Err = 0 Then
        Line Input #lNumArquivo, lRegistro
        If M46V999.gfPreenchido(lRegistro) Then
            If IsNumeric(lRegistro) Then
                If Len(lRegistro) >= 8 Then
                    pCod_Corretor = CLng(Mid$(lRegistro, 2, 7))
                Else
                    lRegistro = Format(lRegistro, "00000000")
                    pCod_Corretor = CLng(Mid$(lRegistro, 2, 7))
                End If
            End If
        End If
    End If
    On Error GoTo 0
    '3. Fecho o arquivo P2842400.EXE
    Close #lNumArquivo
End Sub

Public Function gfForCNPJ(ByVal pCampo As String) As String
'Fun��o: formata CNPJ.

    Dim lAuxiliar As String    'Auxiliar.
    Dim lI As Byte    'Utilizado com For...Next.

    lAuxiliar = ""
    For lI = 1 To Len(pCampo)
        If IsNumeric(Mid(pCampo, lI, 1)) Then
            lAuxiliar = lAuxiliar & Mid(pCampo, lI, 1)
        End If
    Next lI
    If Len(lAuxiliar) = 14 Then
        gfForCNPJ = Left(lAuxiliar, 2) & "." & Mid(lAuxiliar, 3, 3) & "." & Mid(lAuxiliar, 6, 3) & "/" & _
                    Mid(lAuxiliar, 9, 4) & "-" & Right(lAuxiliar, 2)
    Else
        gfForCNPJ = ""
    End If
End Function
Public Function gfForCPF(ByVal pCampo As String) As String
'Fun��o: formata CPF.

    Dim lAuxiliar As String    'Auxiliar.
    Dim lI As Byte    'Utilizado com For...Next.

    lAuxiliar = ""
    For lI = 1 To Len(pCampo)
        If IsNumeric(Mid(pCampo, lI, 1)) Then
            lAuxiliar = lAuxiliar & Mid(pCampo, lI, 1)
        End If
    Next lI
    If Len(lAuxiliar) = 11 Then
        gfForCPF = Left(lAuxiliar, 3) & "." & Mid(lAuxiliar, 4, 3) & "." & Mid(lAuxiliar, 7, 3) & "-" & _
                   Right(lAuxiliar, 2)
    Else
        gfForCPF = ""
    End If
End Function

Public Function gfGraNovPedido(ByVal pBasDados As ADODB.Connection, ByVal pNosNumero As String, _
                               ByVal pNumPedido As String, ByVal pDatVigInicio As String, _
                               ByVal pDatVigTermino As String, ByVal pCodProduto As Integer, _
                               ByVal pCodProtocolo As Integer, ByVal pCodSequencia As Integer, _
                               ByRef pMensagem As String, pCod_Ramo As Integer, _
                               pCodPlano As Long, pCodSegmentacao As Integer, _
                               pobjstcPed As stcA46V702B, tagDynaTrace As String) As Boolean

' Valores padr�o:
' pCod_Ramo = 0
' pCodPlano = 0
' pCodSegmentacao = 0
' pObjstcPed = new stcA46V702B

'Fun��o: grava registro de pedido novo.

'Par�metros de entrada...pBasDados........Base de dados.
'                        pNosNumero.......Nosso n�mero.
'                        pNumPedido.......N�mero do pedido.
'                        pDatVigInicio....Data de vig�ncia: in�cio.
'                        pDatVigTermino...Data de vig�ncia: t�rmino.
'Par�metro de sa�da......pMensagem........Mensagem.

    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "gfGraNovPedido"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Dim lCodCorretor As Long      'C�digo do corretor.
    Dim lInsert As String    'Comando INSERT.
    Dim lRegistro As String    'Registro.
    Dim lObjPed As clsA46V702A    'Objeto da classe clsA46V702A (TAB_PED)
    Dim lObjstcPed As stcA46V702B    'Objeto da classe de estrutura stcA46V702B (TAB_PED)
    Dim lobjCtrlEmis As clsA46V712A    'Objeto da classe de estrutura stcA46V712B (TAB_CTRL_EMIS)
    Dim lObjPedEnd As clsA46V707A    'Objeto da classe de estrutura stcA46V707B (TAB_PED_ENDER)
    Dim lobjCorr As clsA46V705A    'Objeto da classe clsA46V705A (TAB_PED_CORR)
    Dim lObjstcCorr As stcA46V705B    'Objeto da classe de estrutura stcA46V705B (TAB_PED_CORR)
    Dim lobjJuros As clsA46V711A    'Objeto da classe clsA46V711A (TAB_PED_JUROS)
    Dim lQtdRegistro As Long      'Quantidade de registros.
    Dim lObjPedAuto As clsA46V719A    'Objeto da classe de estrutura stcA46V719B (TAB_PED_AUTO)
    Dim lobjEspec As clsA46V726A    'Objeto da classe de estrutura clsA46V726A (TAB_PED_ESPEC)
    Dim lPropTabela As String

    On Error GoTo ControleErro
    gfGraNovPedido = False
    Call M46V999.gfObtSYASMULTINI(lPropTabela)

    '1. Obt�m c�digo do corretor.
    lCodCorretor = M46V999.gCod_Corretor

    '2. Instancio as classes da tabela TAB_PED
    Set lObjPed = New clsA46V702A
    Set lObjPed.M46V999 = M46V999

    Set lObjstcPed = New stcA46V702B
    Set lObjstcPed.M46V999 = M46V999

    Set lObjPed.ConexaoBD = pBasDados

    'Carrego o registro padr�o.
    If pCod_Ramo <> 0 Then
        lObjstcPed.nosso_numero = "PADRAO" & M46V999.gNomUsuario6 & pCod_Ramo
    Else
        lObjstcPed.nosso_numero = "PADRAO" & M46V999.gNomUsuario6
    End If
    If Not lObjPed.mfCarregarColecao(lObjstcPed, lPropTabela, "", True) Then
        pMensagem = lObjPed.UltimoErro
        GoTo ControleErro
    End If

    If lObjPed.Colecao.Count > 0 Then
        Set lObjstcPed = lObjPed.Colecao.Item(1)
    Else
        Call M46V999.mfVerRegPadrao(pMensagem)
    End If

    If Not IsNumeric(pCod_Ramo) Then pCod_Ramo = 0

    'Carrego as informa��es
    With lObjstcPed
        .nosso_numero = pNosNumero
        .Num_Ped = pNumPedido
        .Cod_Ramo = pCod_Ramo
        .Dat_Ini_Vig = pDatVigInicio
        .Dat_Ter_Vig = pDatVigTermino
        .Cod_Tip_Doc = e_TipDoc_Cotacao
        .Cod_Produto = pCodProduto
        .Cod_Protocolo = pCodProtocolo
        .Cod_Sequencia = pCodSequencia
        .Cod_User_Liber = M46V999.gNomUsuario6
        .Cod_Dizer = 0
        .Num_Apol_An = 0
        If M46V111.gNum_Pi = 0 Then
            .NUM_PI = Right(Year(Date), 2) & Format(pNumPedido, "00000000")
        Else
            .NUM_PI = M46V111.gNum_Pi
        End If
        .Cod_Reg = 1
        .Tip_Cust_Emis = 1
        .Val_Cust_Emis = M46V999.gProCusApolice
        .Val_Parc_Min = M46V999.gProPreMinParcela
        If M46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda Then
            .Cod_Segurado = val(M46V999.gCodSeguradoYas)
        End If
        If .Cod_Estip = 0 Or ((.Cod_Estip = 1 Or .Cod_Estip = 2) And .Cod_Produto >= 1080) Then .Cod_Estip = M46V999.gfPlanoBasico(pCod_Ramo)
        If pCodPlano <> 0 Then
            .Cod_Estip = pCodPlano
        End If
        .Cod_Iden_Praz = e_TipoPrazo_ProRata
        .Qtd_Parc = IIf(pCodSegmentacao <> 0 And pCodSegmentacao <> 3 And pCodSegmentacao <> 9 And .Cod_Ramo = 112, 7, 11)
        'Fichas 1200168 e 1200169 Inclus�o dos campos de listagem Renda e Faixa de Faturamento Pessoa f�sica e jur�dica
        .Val_RendaFatu = -1
        .Cod_Segmentacao = pCodSegmentacao
    End With
    '3. Chamo a rotina de inclus�o
    If Not lObjPed.mfIncluir(lObjstcPed, lPropTabela, "", tagDynaTrace) Then
        pMensagem = lObjPed.UltimoErro
        GoTo ControleErro
    End If

    'CTRL_EMIS
    Set lobjCtrlEmis = New clsA46V712A
    Set lobjCtrlEmis.M46V999 = M46V999
    Set lobjCtrlEmis.ConexaoBD = lObjPed.ConexaoBD

    lObjstcPed.objstcCtrlEmis.nosso_numero = pNosNumero
    lObjstcPed.objstcCtrlEmis.Cod_Ramo = pCod_Ramo
    If Not lobjCtrlEmis.mfIncluir(lObjstcPed.objstcCtrlEmis, lPropTabela, "", tagDynaTrace) Then
        pMensagem = lobjCtrlEmis.UltimoErro
        GoTo ControleErro
    End If

    'PED_ENDERECO
    Set lObjPedEnd = New clsA46V707A
    Set lObjPedEnd.M46V999 = M46V999

    Set lObjPedEnd.ConexaoBD = lObjPed.ConexaoBD
    lObjstcPed.objstcPedEndereco.nosso_numero = pNosNumero
    lObjstcPed.objstcPedEndereco.Cod_Ramo = pCod_Ramo
    If Not lObjPedEnd.mfIncluir(lObjstcPed.objstcPedEndereco, lPropTabela, "", tagDynaTrace) Then
        pMensagem = lObjPedEnd.UltimoErro
        GoTo ControleErro
    End If

    'PED_AUTO
    Set lObjPedAuto = New clsA46V719A
    Set lObjPedAuto.M46V999 = M46V999

    Set lObjPedAuto.ConexaoBD = lObjPed.ConexaoBD
    lObjstcPed.objstcPedAuto.nosso_numero = pNosNumero
    lObjstcPed.objstcPedAuto.Cod_Ramo = pCod_Ramo
    If Not lObjPedAuto.mfIncluir(lObjstcPed.objstcPedAuto, lPropTabela, "") Then
        pMensagem = lObjPedEnd.UltimoErro
        GoTo ControleErro
    End If

    'Juros
    'Set lobjJuros = New clsA46V711A
    'Set lobjJuros.M46V999 = M46V999
    'Set lobjJuros.ConexaoBD = lObjPed.ConexaoBD
    'If Not lobjJuros.mfCopiarPadraoParaNovo(pNosNumero) Then
    '    pMensagem = lobjJuros.UltimoErro
    '    GoTo ControleErro
    'End If

    'PED_CORR
    Set lobjCorr = New clsA46V705A
    Set lobjCorr.M46V999 = M46V999

    Set lobjCorr.ConexaoBD = lObjPed.ConexaoBD
    For Each lObjstcCorr In lObjstcPed.ColecaoPedCorr
        lObjstcCorr.nosso_numero = pNosNumero
        lObjstcCorr.Cod_Org = lObjstcPed.Cod_Org
        lObjstcCorr.Cod_Ramo = pCod_Ramo
        lObjstcCorr.Num_Ped = pNumPedido
        '        lObjstcCorr.Cod_Corr = lCodCorretor
        '        lObjstcCorr.Num_Ordem_Corr = 1
        lObjstcCorr.NUM_PI = lObjstcPed.NUM_PI
        Set lObjstcCorr.ObjPai = lObjstcPed
        If Not lobjCorr.mfIncluir(lObjstcCorr, lPropTabela, "", tagDynaTrace) Then
            pMensagem = lobjCorr.UltimoErro
            GoTo ControleErro
        End If
    Next

    'TAB_PED_ESPEC
    Dim lobjstcEspc As stcA46V726B
    Set lobjEspec = New clsA46V726A
    Set lobjEspec.M46V999 = M46V999

    Set lobjEspec.ConexaoBD = lObjPed.ConexaoBD

    For Each lobjstcEspc In lObjstcPed.ColecaoTextoEspec
        lobjstcEspc.nosso_numero = pNosNumero
        lobjstcEspc.Cod_Org = lObjstcPed.Cod_Org
        lobjstcEspc.Cod_Ramo = pCod_Ramo
        lobjstcEspc.Num_Ped = pNumPedido
        lobjstcEspc.NUM_PI = lObjstcPed.NUM_PI
        Set lobjstcEspc.ObjPai = lObjstcPed
        If Not lobjEspec.mfIncluir(lobjstcEspc, lPropTabela, "", tagDynaTrace) Then
            pMensagem = lobjEspec.UltimoErro
            GoTo ControleErro
        End If
    Next

    'TAB_PED_ESPEC2
    'Set lobjEspec = New clsA46V721A
    'Set lobjEspec.M46V999 = M46V999
    'Set lobjEspec.ConexaoBD = lObjPed.ConexaoBD
    'lObjstcPed.objstcEspec.Nosso_Numero = pNosNumero
    'If Not lobjEspec.mfIncluir(lObjstcPed.objstcEspec) Then
    '    pMensagem = lobjEspec.UltimoErro
    '    GoTo ControleErro
    'End If
    Set pobjstcPed = lObjstcPed
    gfGraNovPedido = True
    Set pBasDados = lObjPed.ConexaoBD
    
    
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
    
    '4. Destruo os objetos
    Set lObjPed = Nothing
    Set lobjJuros = Nothing
    Set lObjstcPed = Nothing
    Set lobjCtrlEmis = Nothing
    Set lObjPedEnd = Nothing
    Set lObjPedAuto = Nothing
    Set lobjCorr = Nothing
    Set lObjstcCorr = Nothing
    Set lobjEspec = Nothing
    Set lobjstcEspc = Nothing
    Exit Function
ControleErro:
    Call M46V999.gpGraLog(1, "gfGraNovPedido ERRO - " & Err & " " & Err.Description & " " & pMensagem)
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "gfGraNovPedido ERRO - " & Err & " " & Err.Description & " " & pMensagem, 2, Class_Name, methodName)
    Set lObjPed = Nothing
    Set lobjJuros = Nothing
    Set lObjstcPed = Nothing
    Set lobjCtrlEmis = Nothing
    Set lObjPedEnd = Nothing
    Set lObjPedAuto = Nothing
    Set lobjCorr = Nothing
    Set lObjstcCorr = Nothing
    Set lobjEspec = Nothing
    Exit Function
    Resume
End Function
Public Function gfGraRenYasuda(ByVal pNumApolice As String, ByVal pIteApolice As String, _
                               ByRef pNosNumero As String, _
                               ByRef pMensagem As String, ByRef pMsgAlerta As String, tagDynaTrace As String) As Boolean
'Fun��o: grava registros de cota��o/proposta de renova��o Yasuda.

'Par�metros de entrada...pNumApolice...N�mero da ap�lice.
'                        pIteApolice...N�mero do item da ap�lice.
'                        pNosNumero....Nosso n�mero.
'Par�metro de sa�da......pMensagem.....Mensagem.

    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "gfGraRenYasuda"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Dim bdP00APOL As ADODB.Connection    'Base de dados: P00Apol.
    Dim lCodProduto As Integer      'C�digo do produto.
    Dim lCodProtocolo As Long           'C�digo do protocolo.
    Dim lCodSequencia As Long           'C�digo da seq��ncia.
    Dim lDescTarifa As String       'Descri��o da tarifa.
    Dim lNosNumero As String      'Nosso n�mero.
    Dim lNumPedido As String      'N�mero do pedido.
    Dim lQuaIteComErro As Integer         'Quantidade de itens com erro.
    Dim lQuaIteSemErro As Integer         'Quantidade de itens sem erro.
    Dim lVigDatIni As Long        'Data de vig�ncia: in�cio.
    Dim lVigDatTer As Long        'Data de vig�ncia: t�rmino.
    Dim rsP00APOL As clsYasRecordSet    'Registro: P00Apol.
    Dim rsRet As clsYasRecordSet    'Registro: P00Apol.
    Dim lNumEndo As Long
    Dim lRegistro As String
    Dim lNomeTab As String
    Dim lNumItem As Integer
    Dim lAchouItem As Boolean
    Dim rdsTemp As clsYasRecordSet    '1100177 - Renova��o Apolice LMI
    Dim lNumCotac_LMI As String         '1100177 - Renova��o Apolice LMI
    Dim lSelect As String
    Dim lStcPed As stcA46V702B

    If IsNumeric(pMensagem) Then
        lCodProduto = pMensagem
    End If
    pMensagem = Empty
    
    M46V111.gCod_Dizer = 0
    On Error GoTo ControleErro
    gfGraRenYasuda = False

    If M46V999.gPrograma = "D046O003" Then
        lNomeTab = "Syasportal.dbo.TAB_APOL"
    Else
'        If M46V999.gBaseLocal = True Then
            lNomeTab = "TAB_APOL"
'        Else
'            lNomeTab = "GED.DBO.TAB_PED"
'        End If
    End If
    pMsgAlerta = ""
    lNumCotac_LMI = ""    '1100177 - Renova��o Apolice LMI

BuscaApolice:
    '1. Obt�m registro da ap�lice.
    'a) Monta comando SELECT.
    lSelect = "SELECT distinct "
    lSelect = lSelect & lNomeTab & ".*, " & lNomeTab & "_LOC.NUM_ITEM, COD_TIP_LOGR_RISCO, "
    lSelect = lSelect & " NOM_LOGR_RISCO, "
'    If (M46V999.gBaseLocal = True Or M46V999.gPrograma = "D046O003") And lNumCotac_LMI = "" Then
'        lSelect = lSelect & " Dsc_Compl_Risco , "
'    Else
        lSelect = lSelect & " left('N.' + (select cast(NUM_LOGR_RISCO as varchar) from " & lNomeTab & "_Insp WITH (NOLOCK) where "
'        If lNumCotac_LMI = "" Then    '1100177 - Renova��o Apolice LMI
            lSelect = lSelect & lNomeTab & "_Insp.Num_PI = " & lNomeTab & "_loc.num_pi and " & lNomeTab & "_Insp.num_ped = "
'        Else
'            lSelect = lSelect & lNomeTab & "_Insp.NOSSO_NUMERO = " & lNomeTab & "_loc.NOSSO_NUMERO and " & lNomeTab & "_Insp.num_ped = "
'        End If
        lSelect = lSelect & lNomeTab & "_loc.num_ped and " & lNomeTab & "_Insp.num_item = " & lNomeTab & "_loc.num_item) +' - '+  rtrim(DSC_COMPL_RISCO) ,30) as Dsc_Compl_Risco,"
'    End If
    lSelect = lSelect & " Nom_Bairro_Risco , "
    lSelect = lSelect & lNomeTab & "_LOC.NUM_PI as NUM_PI_Item, "  '1100177 - Renova��o Apolice LMI
    lSelect = lSelect & " NOM_CID_RISCO, SIG_UF_RISCO, NUM_CEP_RISCO, "
    lSelect = lSelect & " PER_DESC_APAR, PER_DESC_PROGRES, PER_IS_VR, "
    lSelect = lSelect & " VAL_COEF_IS_VR, VAL_VR, TAB_CLASF, COD_CLASF, "
    lSelect = lSelect & " COD_COMPL_CLASF, DSC_CLASF, COD_CLASSE_OCUP, "
    lSelect = lSelect & " COD_CONSTR,  COD_SIT_ITEM, PER_BONUS, "
    lSelect = lSelect & " QTD_EMPREG, NUM_INSP, COD_IDEN_SEGUR_01, "
    lSelect = lSelect & " COD_IDEN_SEGUR_02, COD_IDEN_SEGUR_03, COD_IDEN_SEGUR_04, "
    lSelect = lSelect & " COD_IDEN_SEGUR_05, COD_IDEN_SEGUR_06, COD_IDEN_SEGUR_07, "
    lSelect = lSelect & " COD_IDEN_SEGUR_08, COD_IDEN_SEGUR_09, COD_IDEN_SEGUR_10, "
    lSelect = lSelect & " PER_TOT_SEGUR, COD_IDEN_PROTEC_01, COD_IDEN_PROTEC_02, "
    lSelect = lSelect & " COD_IDEN_PROTEC_03, COD_IDEN_PROTEC_04, COD_IDEN_PROTEC_05, "
    lSelect = lSelect & " COD_IDEN_PROTEC_06, COD_IDEN_PROTEC_07, COD_IDEN_PROTEC_08, "
    lSelect = lSelect & " COD_IDEN_PROTEC_09, COD_IDEN_PROTEC_10, PER_TOT_PROTEC, "
    lSelect = lSelect & " TIP_COBERT, "

    If M46V999.gBaseLocal Then
        lSelect = lSelect & " Ind_Sin_Inspec, "
    Else
        lSelect = lSelect & " (select count(v.num_apol) from RamosDiversos..tb_aviso_sin_diversos AS a WITH (NOLOCK) inner join Ged..TB_Valores_Socket as V WITH (NOLOCK)"
        lSelect = lSelect & "     on a.num_apol = v.num_apol and  a.num_item  = v.num_item  inner join  RamosDiversos.dbo.Tbm_Lim_Sinistro as L WITH (NOLOCK)"
        lSelect = lSelect & "     on l.Cod_cobert = cod_cobert0 and l.cod_ramo = v.Cod_ramo"
        lSelect = lSelect & "     where tip_pagto in (211,212,213) and dat_contab > 0 and  iden_estor = 0 and Cod_Tarif_Fim = 9999 and val_estim_sin > Val_Max_Sin and"
        lSelect = lSelect & "     v.num_apol = " & pNumApolice & " and v.num_item =  " & lNomeTab & "_LOC.Num_Item) as Ind_Sin_Inspec , "
    End If
    lSelect = lSelect & " " & lNomeTab & "_LOC.COD_PRODUTO, COD_REGIAO, COD_IDEN_TARIF, "
    lSelect = lSelect & " " & lNomeTab & "_LOC.NUM_ENDO AS NUM_ENDO2 "


    lSelect = lSelect & " , LCompl.Ind_VerbaSep , LCompl.Val_Vr_Lc, LCompl.Val_LMG_Lc, LCompl.Per_IS_Vr_Lc , LCompl.Coef_IS_Vr_Lc "
    lSelect = lSelect & " ,Tip_Plano_Vida"
    lSelect = lSelect & " ,Cod_Plano_Cobertura"
    lSelect = lSelect & " ,Qtd_Vidas"
    lSelect = lSelect & " ,Val_LMG_Vida"
    lSelect = lSelect & " ,Ind_Aux_Funeral"
    lSelect = lSelect & " ,Val_LMG_Funeral"
    lSelect = lSelect & " ,Ind_Inc_Conjuge"
    lSelect = lSelect & " ,Val_Aux_Cesta"
    lSelect = lSelect & " ,POSSUI_ELEVADOR"
    lSelect = lSelect & " ,BonusAlterado"
    lSelect = lSelect & " ,IND_SINISTRO"
    lSelect = lSelect & " ,(SELECT top 1 Per_Desc_LMI FROM RamosDiversos.DBO.TAB_PED A WITH (NOLOCK) WHERE " & lNomeTab & "_LOC.NUM_APOL = A.NUM_APOL AND " & lNomeTab & "_LOC.NUM_ENDO = A.NUM_ENDO) AS Per_Desc_LMI"

    If (M46V999.gPrograma = "D046O003") Then
        lSelect = lSelect & " FROM " & lNomeTab & "_CORR WITH (NOLOCK) INNER JOIN " & lNomeTab & " WITH (NOLOCK) ON "
        lSelect = lSelect & " " & lNomeTab & "_CORR.NUM_APOL = " & lNomeTab & ".NUM_APOL AND "
        lSelect = lSelect & " " & lNomeTab & "_CORR.NUM_ENDO = " & lNomeTab & ".NUM_ENDO AND "
        lSelect = lSelect & " Cod_corr in (SELECT * FROM [Ged].[dbo].[Func_BuscaCorrGrupoAtivo] ('" & Format(M46V999.gCod_Corretor, "0000000") & "', 112))"
    Else
        lSelect = lSelect & " FROM " & lNomeTab & " WITH (NOLOCK)"
    End If

    lSelect = lSelect & " INNER JOIN " & lNomeTab & "_LOC WITH (NOLOCK) ON "
    'lSelect = lSelect & " " & lNomeTab & "_LOC.NUM_APOL = " & lNomeTab & ".NUM_APOL "
    If lNumCotac_LMI = "" Then    '1100177 - Renova��o Apolice LMI
        lSelect = lSelect & " " & lNomeTab & "_LOC.NUM_APOL = " & lNomeTab & ".NUM_APOL "
    Else
        lSelect = lSelect & " " & lNomeTab & "_LOC.NUM_PI = " & lNomeTab & ".NUM_PI "
    End If
    
'    If (Not M46V999.gBaseLocal) And (Not M46V999.gPrograma = "D046O003") Then
'        lSelect = lSelect & " AND " & lNomeTab & "_LOC.NUM_ENDO = " & lNomeTab & ".NUM_ENDO "
'    End If

    lSelect = lSelect & " LEFT JOIN " & lNomeTab & "_LOC_COMPL AS LCompl WITH (NOLOCK) ON "
'    lSelect = lSelect & " LEFT JOIN  RAMOSDIVERSOS.DBO.Tab_Ped_Loc_Compl AS LCompl ON "

    lSelect = lSelect & " " & lNomeTab & "_LOC.NUM_APOL = LCompl.NUM_APOL and  "
    lSelect = lSelect & " " & lNomeTab & "_LOC.NUM_Item = LCompl.NUM_Item "
    If (Not M46V999.gBaseLocal) And (Not M46V999.gPrograma = "D046O003") Then
        lSelect = lSelect & " AND " & lNomeTab & "_LOC.NUM_ENDO = LCompl.NUM_ENDO "
    End If
    
'    lSelect = lSelect & " INNER JOIN [GED]..[TAB_CTRL_EMIS] AS [TCE] ON "

    If lNumCotac_LMI = "" Then    '1100177 - Renova��o Apolice LMI
        lSelect = lSelect & " Where " & lNomeTab & "_LOC.Num_Apol = " & pNumApolice
    Else
        lSelect = lSelect & " AND   " & lNomeTab & "_LOC.NOSSO_NUMERO = " & lNomeTab & ".NOSSO_NUMERO"
        lSelect = lSelect & " Where " & lNomeTab & ".NOSSO_NUMERO = " & M46V999.gfFormatarCamposSQL(lNumCotac_LMI)
    End If

    'Ficha 1000782 - N�o permitir Renova��o LMI para Corretores
    If M46V999.gTipUsuario = e_Tipo_Usu_Corretor Then
        lSelect = lSelect & " AND " & lNomeTab & "_LOC.NOM_LOGR_RISCO not like '%CONFORME ESPEC%'"
    End If
    lSelect = lSelect & " AND [COD_DIZER] NOT IN (4,10)          "
    If val(pIteApolice) <> 0 Then
        lSelect = lSelect & " AND " & lNomeTab & "_LOC.NUM_ITEM = " & pIteApolice
    End If
    
    ' If Not M46V999.gBaseLocal Then
    '     lSelect = lSelect & " and " & lNomeTab & ".NUM_PI IN (SELECT NUM_PI "
    '     lSelect = lSelect & " From [tab_ctrl_emis]"
    '     lSelect = lSelect & " WHERE  [cod_iden_stat_acao] = 90 ) "
    ' End If
    
    'lSelect = lSelect & " ORDER BY " & lNomeTab & ".NUM_ENDO desc ,NUM_ITEM"
    lSelect = lSelect & " ORDER BY " & lNomeTab & "_LOC.NUM_ITEM, " & lNomeTab & ".DAT_EMIS desc"

    'b) Abre arquivo.
    If M46V999.gfAbrBasDados("P00APOL", M46V999.gTipP00APOL, M46V999.gEndP00APOL, bdP00APOL, pMensagem) = False Then
        Exit Function
    End If
    'c) Obt�m registro.
    If M46V999.gfObtRegistro(bdP00APOL, lSelect, rsP00APOL, pMensagem) = False Then
        Call M46V999.gpFecha2(bdP00APOL)
        Exit Function
    End If
    lAchouItem = True
    'd) N�o encontrou registro de ap�lice.
    If rsP00APOL.EOF = True Then
        '1001106 - Renova��o Roubo
        lSelect = "SELECT * FROM " & lNomeTab & " WITH (NOLOCK)"
        lSelect = lSelect & " Where Num_Apol = " & pNumApolice
        lSelect = lSelect & " AND NUM_ENDO = 0"
        
        If M46V999.gfObtRegistro(bdP00APOL, lSelect, rsP00APOL, pMensagem) = False Then
            Call M46V999.gpFecha2(bdP00APOL)
            Exit Function
        End If
        If rsP00APOL.EOF = True Then
            '===1) Se n�o encontrou registro, verifica se tem ap�lice.
            If M46V999.gfPreenchido(pIteApolice) Then
                lSelect = "SELECT * FROM " & lNomeTab & "_LOC WITH (NOLOCK) "
                lSelect = lSelect & " Where Num_Apol = " & pNumApolice
                lSelect = lSelect & " order by NUM_ENDO desc "
                
                If M46V999.gfObtRegistro(bdP00APOL, lSelect, rsP00APOL, pMensagem) = False Then
                    Call M46V999.gpFecha2(bdP00APOL)
                    Exit Function
                End If
            End If
            '*Luiz Alberto em 15/01/2016
            If rsP00APOL.EOF = True Then
Verifica:
                lSelect = ""
                lSelect = "SELECT [TPD].[NUM_APOL] "
                lSelect = lSelect & " FROM [GED].[DBO].[TAB_PED]  AS [TPD] WITH(NOLOCK) "
                lSelect = lSelect & " INNER JOIN [GED]..[TAB_CTRL_EMIS] AS [TCE] WITH (NOLOCK) "
                lSelect = lSelect & " ON [TPD].[NUM_APOL]   = [TCE].[NUM_APOL] "
                lSelect = lSelect & " AND [TPD].[NUM_ENDO]  = [TCE].[NUM_ENDO] "
                lSelect = lSelect & " AND [TPD].[COD_RAMO]  = [TCE].[COD_RAMO] "
                lSelect = lSelect & " Where [TPD].[Num_Apol] <> 0 "
                lSelect = lSelect & " AND [TPD].[COD_RAMO]  IN (112,113,114) "
                lSelect = lSelect & " AND [TCE].[COD_IDEN_STAT_ACAO] = 70 "
                lSelect = lSelect & " AND [TPD].[NUM_APOL]  = " & pNumApolice
                'e)Obt�m registro.
                If M46V999.gfObtRegistro(bdP00APOL, lSelect, rsP00APOL, pMensagem) = False Then
                    Call M46V999.gpFecha2(bdP00APOL)
                    Exit Function
                End If
                If rsP00APOL.EOF = True Then
                        pMensagem = "FNL0012 - Ap�lice " & pNumApolice & " n�o localizada."
                    Else
                        pMensagem = "FNL0012 - Ap�lice " & pNumApolice & " Cancelada - N�o dispon�vel para renova��o."
                                End If
                        Else
                pMensagem = "FNL0013 - Item  " & pIteApolice & " n�o localizado na ap�lice " & _
                            pNumApolice & "."
            End If
            Call M46V999.gpFecha2(bdP00APOL)
            Exit Function
        Else
            If rsP00APOL("cod_ramo") <> 150 And rsP00APOL("cod_ramo") <> 710 Then
                GoTo Verifica
            End If
            lAchouItem = False
        End If
    End If

    If rsP00APOL("NOM_LOGR_RISCO") Like "*CONFORME ESPEC*" Then

        If rsP00APOL("num_endo") <> 0 Then
            If M46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then
                pMsgAlerta = "Aten��o - Ap�lice com Endosso n� " & rsP00APOL("NUM_ENDO") & " - Verificar!"
            Else
                pMsgAlerta = "Dados carregados conforme emiss�o inicial da ap�lice. Caso necess�rio, atualize as informa��es antes de transmitir."
            End If
        End If
        GoTo BuscaPorEndosso
    End If


    '1100177 - Renova��o Apolice LMI
    If (M46V999.gTipUsuario = e_Tipo_Usu_Yasuda) And (M46V999.gBaseLocal = False) And (lNumCotac_LMI = "") Then
        If Left(rsP00APOL("NOM_LOGR_RISCO"), 14) = "CONFORME ESPEC" Then
BuscaPorEndosso:
            lSelect = "Select NUM_COTAC_LMI FROM RAMOSDIVERSOS.DBO.TAB_PED WITH (NOLOCK)"
            lSelect = lSelect & " Where NUM_APOL = " & rsP00APOL("NUM_APOL")
            If M46V999.gfObtRegistro(bdP00APOL, lSelect, rdsTemp, pMensagem) = False Then
                Call M46V999.gpFecha2(bdP00APOL)
                Exit Function
            End If
            If rdsTemp.EOF Then
                If rsP00APOL.EOF = False Then
                    Call M46V999.gpGraLog(1, "Aten��o - Ap�lice com Endosso n� " & rsP00APOL("NUM_ENDO") & " - Verificar!")
                    rsP00APOL.MoveNext
                    If rsP00APOL.EOF Then GoTo Verifica
                    GoTo BuscaPorEndosso
                Else
                    GoTo Verifica
                End If
            ElseIf Trim(M46V999.gfTratarNulo(rdsTemp("NUM_COTAC_LMI"), tpoString)) <> "" Then
                lNumCotac_LMI = Trim(M46V999.gfTratarNulo(rdsTemp("NUM_COTAC_LMI"), tpoString))
                lNomeTab = "RAMOSDIVERSOS.DBO.TAB_PED"
                GoTo BuscaApolice
            End If
        End If
    End If

    '3. Posiciona nosso n�mero.
    If Not M46V999.gfPreenchido(pNosNumero) Then
        If M46V999.gfCalNosNumero(lNosNumero, pMensagem) = False Then
            Call M46V999.gpFecha2(bdP00APOL)
            Exit Function
        End If
        pNosNumero = lNosNumero
    Else
        lNosNumero = pNosNumero
    End If

    '4. Posiciona n�mero do pedido.
    If M46V999.gfCalNumPedido(lNumPedido, pMensagem) = False Then
        Call M46V999.gpFecha2(bdP00APOL)
        Exit Function
    End If
    M46V999.gPrefixo = "Cota��o " & lNumPedido & ". "

    '5. Posiciona vig�ncia.
    rsP00APOL.MoveFirst

    'AJUSTE PARA DESCONSIDERAR A DATA DE TERMINO DE VIGENCIAS DOS ITENS DE LMI ANTIGO E CONSIDERAR O DA APOLICE
    If rsP00APOL!Dat_Ter_Vig > M46V999.gfForData(1, Format(Date + 100, "dd/mm/yyyy")) And lNumCotac_LMI = "" Then
        'CRQ000000458062 - Alterar de 60 para 100 dias
        pMensagem = "FGE0088 - Ap�lice n�o dispon�vel para renova��o. T�rmino de vig�ncia em " & M46V999.gfForData(2, rsP00APOL!Dat_Ter_Vig) & "."
        Exit Function
    End If

    lVigDatIni = rsP00APOL!Dat_Ter_Vig
    lVigDatTer = lVigDatIni + 10000
    
    'Posicionar o mesmo plano
    'sandra 900218
    M46V111.gCod_plano = rsP00APOL!Cod_Estip

    M46V999.gCod_Ramo = rsP00APOL!Cod_Ramo

    If M46V111.gCod_plano >= 1001 And M46V111.gCod_plano <= 1010 Then
        'De/Para dos planos da Mercedes
        Select Case M46V111.gCod_plano
        Case 1001
            M46V111.gCod_plano = 2021
        Case 1002
            M46V111.gCod_plano = 2022
        Case 1003
            M46V111.gCod_plano = 2023
        Case 1004
            M46V111.gCod_plano = 2024
        Case 1005
            M46V111.gCod_plano = 2025
        Case 1006
            M46V111.gCod_plano = 2026
        Case 1007
            M46V111.gCod_plano = 2027
        Case 1008
            M46V111.gCod_plano = 2028
        Case 1009
            M46V111.gCod_plano = 2029
        Case 1010
            M46V111.gCod_plano = 2030
        End Select

        M46V999.gProCusApolice = 20
        M46V999.gProPreMinParcela = 0
    End If
    
    'Argumento passado na chamada do EXE se for igual a DEV � a vers�o de homologa��o
'    If InStr(UCase(Command), "DEV") = 0 And M46V999.gPrograma <> "D046O003" Then
'        If rsP00APOL!Cod_Org <> 2 And M46V999.gTarifa_Apolice < 1080 Then
'            lCodProduto = 1076
'        End If
'    End If

    '6. Posiciona produto que atenda � vig�ncia.
    If Not M46V999.gfBusProduto(Format(Date, "YYYYMMDD"), lCodProduto, lCodProtocolo, lCodSequencia, lDescTarifa, pMensagem, rsP00APOL("Cod_Ramo")) Then
        Call M46V999.gpFecha2(bdP00APOL)
        Exit Function
    End If
    If lCodProduto = 0 Then
        pMensagem = "FGE0088 - N�o h� produto (tarifa) dispon�vel para c�lculo da renova��o da " & _
                    "ap�lice " & pNumApolice & ", de acordo com a data de in�cio de vig�ncia (" & _
                    M46V999.gfForData(2, lVigDatIni) & ")." & vbLf & vbLf & "Aguardar novo produto (tarifa)." & _
                    vbLf & vbLf & "Caso seja necess�rio, o c�lculo deve ser feito pela Yasuda."
        Exit Function
    End If

    'Verifica Black List
    If M46V999.gPrograma = "D046O003" Then
        Dim lBlackList As Boolean
        pMensagem = Empty
        
        lBlackList = M46V999.Is_BlackList(lCodProduto, rsP00APOL!Cod_Ramo, rsP00APOL!Tip_Pessoa, rsP00APOL!Num_Cgc_Cpf, pMensagem)
        
        If M46V999.gfPreenchido(pMensagem) Then
            Exit Function
        Else
            If lBlackList Then
                pMensagem = "Por motivos t�cnicos n�o � poss�vel prosseguir com esta renova��o."
                Exit Function
            End If
        End If
    End If

    'Verifica as regras para enquadrar o plano Residencial Supremo
    M46V999.gCodProduto = lCodProduto
    
    If M46V999.LiberaPlanoResidencialSupremo(30, e_TipEmissao_RenYasuda) Then
        Call M46V999.ContemResidencialSupremo(rsP00APOL)
        
        '' 2017-03-10: Corrigir bug Supremo/RsEOF
        If Not rsP00APOL Is Nothing And rsP00APOL.RecordCount > 0 Then
            rsP00APOL.MoveFirst
        End If
    End If
    
    '7. Grava registro de renova��o Yasuda.
    If M46V999.gCod_Ramo = e_CodRamo_Riscos_diversos Then
        If mfGraRenComum(rsP00APOL, lNosNumero, CLng(lNumPedido), lVigDatIni, lVigDatTer, lCodProduto, _
                         CInt(lCodProtocolo), CInt(lCodSequencia), 0, bdP00APOL, lNumCotac_LMI, pNumApolice, pMensagem, lStcPed, tagDynaTrace) = False Then
            Call M46V999.gpFecha2(bdP00APOL)
            Exit Function
        End If
    Else
        If mfGraRenComum(rsP00APOL, lNosNumero, CLng(lNumPedido), lVigDatIni, lVigDatTer, lCodProduto, _
                         CInt(lCodProtocolo), CInt(lCodSequencia), 0, bdP00APOL, lNumCotac_LMI, pNumApolice, pMensagem, lStcPed, tagDynaTrace) = False Then
            Call M46V999.gpFecha2(bdP00APOL)
            Exit Function
        End If
    End If

    '8. Grava registros de itens, coberturas e descontos.
    rsP00APOL.MoveLast      'Posiciona quantidade de registros.
    rsP00APOL.MoveFirst
    lNumEndo = 0
    lNumItem = 0
    
    While Not rsP00APOL.EOF
        '1001106 - Renova��o Roubo
        If (rsP00APOL("COD_RAMO") = e_CodRamo_Roubos Or rsP00APOL("COD_RAMO") = e_CodRamo_Riscos_diversos) And lAchouItem = False Then
            If Not mfGraRenItem_Roubo(rsP00APOL, lNosNumero, bdP00APOL, pMensagem, CLng(lNumPedido), 0, tagDynaTrace) Then
                Call M46V999.gpGraLog(1, pMensagem)
                Exit Function
            End If
        Else
            If lNumItem <> rsP00APOL("Num_Item") Then
                
                M46V999.gNum_Apolice = lStcPed.Num_Apol_An
                If M46V999.gTarifa_Apolice < 1080 And M46V999.gCodProduto >= 1080 And rsP00APOL!Cod_Constr = 2 And lStcPed.objstcCtrlEmis.cod_origem <> 2 Then
                    rsP00APOL!Cod_Constr = 4
                End If

                If Not mfGraRenItem(rsP00APOL, lNosNumero, bdP00APOL, _
                                    pMensagem, CLng(lNumPedido), M46V111.gCod_Dizer, lNumCotac_LMI, pNumApolice, lStcPed, tagDynaTrace) Then
                    Call M46V999.gpGraLog(1, pMensagem)
                    Exit Function
                End If
                lNumItem = rsP00APOL("Num_Item")
            End If
        End If
        rsP00APOL.MoveNext
    Wend
    Call M46V999.gpFecha2(bdP00APOL)

    '9. Grava registro de comunica��o.
    If M46V111.gfGraRegComunicacao(lNosNumero, pMensagem) = False Then
        Call M46V999.gpGraLog(1, pMensagem)
        Exit Function
    End If

    Set bdP00APOL = Nothing
    Set rsP00APOL = Nothing
    Set rsRet = Nothing
    Set rdsTemp = Nothing
    Set lStcPed = Nothing

    gfGraRenYasuda = True
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
    Exit Function
ControleErro:
    Set bdP00APOL = Nothing
    Set rsP00APOL = Nothing
    Set rsRet = Nothing
    Set rdsTemp = Nothing
    Set lStcPed = Nothing
    Call M46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "gfGraRenYasuda ERRO - " & Err & " " & Err.Description & " " & pMensagem, 2, Class_Name, methodName)
    Call M46V999.gpGraLog(0, "gfGraRenYasuda ERRO - " & Err & " " & Err.Description & " " & pMensagem)
End Function

Public Function gfGraEndossoYasuda(ByVal pNumApolice As String, ByVal pIteApolice As String, _
                                   ByRef pNosNumero As String, pIndSinistro As e_SimNao, _
                                   ByVal pCod_Dizer As Integer, ByRef pMensagem As String, _
                                   pCadastraComum As Boolean, pStcPed As stcA46V702B, tagDynaTrace As String) As Boolean
'Fun��o: grava registros de cota��o/proposta de Endosso Yasuda.

'Par�metros de entrada...pNumApolice...N�mero da ap�lice.
'                        pIteApolice...N�mero do item da ap�lice.
'                        pNosNumero....Nosso n�mero.
'Par�metro de sa�da......pMensagem.....Mensagem.

    Dim bdP00APOL As ADODB.Connection    'Base de dados: P00Apol.
    Dim bdP00MPED As ADODB.Connection    'Base de dados: P00MPED.
    Dim lCodProduto As Integer  'C�digo do produto.
    Dim lCodProtocolo As Long       'C�digo do protocolo.
    Dim lCodSequencia As Long       'C�digo da seq��ncia.
    Dim lDescTarifa As String   'Descri��o da tarifa.
    Dim lNosNumero As String  'Nosso n�mero.
    Dim lNumPedido As String  'N�mero do pedido.
    Dim lQuaIteComErro As Integer     'Quantidade de itens com erro.
    Dim lQuaIteSemErro As Integer     'Quantidade de itens sem erro.
    Dim lVigDatIni As Long    'Data de vig�ncia: in�cio.
    Dim lVigDatTer As Long    'Data de vig�ncia: t�rmino.
    Dim rsP00APOL As clsYasRecordSet    'Registro: P00Apol.
    Dim rsRet As clsYasRecordSet
    Dim lNumEndo As Long
    Dim lRegistro As String
    Dim lNomeTab As String
    Dim lNumItem As Integer
    Dim lSelect As String
    Dim lIndExclusao As Boolean
    Dim lStcPed As stcA46V702B

    On Error GoTo ControleErro
    gfGraEndossoYasuda = False


    If M46V999.gPrograma = "D046O003" Then
        lNomeTab = "Syasportal.dbo.TAB_APOL"
    Else
        'If M46V999.gBaseLocal = True Then
        lNomeTab = "TAB_APOL"
        'Else
        '    lNomeTab = "GED.DBO.TAB_PED"
        'End If
    End If

    If val(pNumApolice) = 0 Then
        pMensagem = "FNL0012 - Ap�lice " & pNumApolice & " zerada."
        Exit Function
    End If

    If val(pIteApolice) = 0 And (pCod_Dizer = e_CodDizerexclusao Or pCod_Dizer = e_CodDizerAlteracao) Then
        pMensagem = "FNL0012 - Numero do Item " & pIteApolice & " zerado."
        Exit Function
    End If

    'Se for endosso de exclus�o, eu testo o cancelamento, se tiver mais de 1 item ent�o
    'continuo como exclus�o, caso contr�rio, calcula como cancelamento mesmo
    lIndExclusao = False
    If pCod_Dizer = 4 And pCadastraComum = True Then
        lIndExclusao = True
        pCod_Dizer = 10
    End If




    '1. Obt�m registro da ap�lice.
    'a) Monta comando SELECT.

    'lSelect = "SELECT distinct "
    'lSelect = lSelect & " " & lNomeTab & ".*, NUM_ITEM, COD_TIP_LOGR_RISCO, "
    'lSelect = lSelect & " NOM_LOGR_RISCO, DSC_COMPL_RISCO, NOM_BAIRRO_RISCO, "
    'lSelect = lSelect & " NOM_CID_RISCO, SIG_UF_RISCO, NUM_CEP_RISCO, "
    'lSelect = lSelect & " PER_DESC_APAR, PER_DESC_PROGRES, PER_IS_VR, "
    'lSelect = lSelect & " VAL_COEF_IS_VR, VAL_VR, TAB_CLASF, COD_CLASF, "
    'lSelect = lSelect & " COD_COMPL_CLASF, DSC_CLASF, COD_CLASSE_OCUP, "
    'lSelect = lSelect & " COD_CONSTR,  COD_SIT_ITEM, PER_BONUS, "
    'lSelect = lSelect & " QTD_EMPREG, NUM_INSP, COD_IDEN_SEGUR_01, "
    'lSelect = lSelect & " COD_IDEN_SEGUR_02, COD_IDEN_SEGUR_03, COD_IDEN_SEGUR_04, "
    'lSelect = lSelect & " COD_IDEN_SEGUR_05, COD_IDEN_SEGUR_06, COD_IDEN_SEGUR_07, "
    'lSelect = lSelect & " COD_IDEN_SEGUR_08, COD_IDEN_SEGUR_09, COD_IDEN_SEGUR_10, "
    'lSelect = lSelect & " PER_TOT_SEGUR, COD_IDEN_PROTEC_01, COD_IDEN_PROTEC_02, "
    'lSelect = lSelect & " COD_IDEN_PROTEC_03, COD_IDEN_PROTEC_04, COD_IDEN_PROTEC_05, "
    'lSelect = lSelect & " COD_IDEN_PROTEC_06, COD_IDEN_PROTEC_07, COD_IDEN_PROTEC_08, "
    'lSelect = lSelect & " COD_IDEN_PROTEC_09, COD_IDEN_PROTEC_10, PER_TOT_PROTEC, "
    'lSelect = lSelect & " " & lNomeTab & "_LOC.COD_PRODUTO, COD_REGIAO, COD_IDEN_TARIF"
    'lSelect = lSelect & " FROM " & lNomeTab & "_LOC INNER JOIN " & lNomeTab & " ON "
    'lSelect = lSelect & " " & lNomeTab & "_LOC.NUM_APOL = " & lNomeTab & ".NUM_APOL "
    'lSelect = lSelect & " Where " & lNomeTab & "_LOC.Num_Apol = " & pNumApolice
    'If Val(pIteApolice) <> 0 Then
    '    lSelect = lSelect & " AND NUM_ITEM = " & pIteApolice
    'End If
    'lSelect = lSelect & " ORDER BY " & lNomeTab & ".NUM_ENDO desc ,NUM_ITEM"

BuscaApolice:


    lSelect = "SELECT distinct "
    lSelect = lSelect & lNomeTab & ".*, " & lNomeTab & "_LOC.NUM_ITEM, COD_TIP_LOGR_RISCO, "
    lSelect = lSelect & " NOM_LOGR_RISCO, "
    lSelect = lSelect & " Dsc_Compl_Risco , "
    If M46V999.gBaseLocal = True Then
        lSelect = lSelect & " 0 as Num_Logr_Risco, "
    Else
        'lSelect = lSelect & " left('N.' + (select cast(NUM_LOGR_RISCO as varchar) from " & lNomeTab & "_Insp where "
        'lSelect = lSelect & lNomeTab & "_Insp.Num_PI = " & lNomeTab & "_loc.num_pi and " & lNomeTab & "_Insp.num_ped = "
        'lSelect = lSelect & lNomeTab & "_loc.num_ped and " & lNomeTab & "_Insp.num_item = " & lNomeTab & "_loc.num_item) +' - '+  rtrim(DSC_COMPL_RISCO) ,30) as Dsc_Compl_Risco,"
        lSelect = lSelect & " (select NUM_LOGR_RISCO  from " & lNomeTab & "_Insp WITH (NOLOCK) where "
        lSelect = lSelect & lNomeTab & "_Insp.Num_PI = " & lNomeTab & "_loc.num_pi and " & lNomeTab & "_Insp.num_ped = "
        lSelect = lSelect & lNomeTab & "_loc.num_ped and " & lNomeTab & "_Insp.num_item = " & lNomeTab & "_loc.num_item)  as NUM_LOGR_RISCO ,"

    End If
    lSelect = lSelect & " Nom_Bairro_Risco , "
    lSelect = lSelect & lNomeTab & "_LOC.NUM_PI as NUM_PI_Item, "  '1100177 - Renova��o Apolice LMI
    lSelect = lSelect & " NOM_CID_RISCO, SIG_UF_RISCO, NUM_CEP_RISCO, "
    lSelect = lSelect & " PER_DESC_APAR, PER_DESC_PROGRES, PER_IS_VR, "
    lSelect = lSelect & " VAL_COEF_IS_VR, VAL_VR, TAB_CLASF, COD_CLASF, "
    lSelect = lSelect & " COD_COMPL_CLASF, DSC_CLASF, COD_CLASSE_OCUP, "
    lSelect = lSelect & " COD_CONSTR,  COD_SIT_ITEM, PER_BONUS, "
    lSelect = lSelect & " QTD_EMPREG, NUM_INSP, COD_IDEN_SEGUR_01, "
    lSelect = lSelect & " COD_IDEN_SEGUR_02, COD_IDEN_SEGUR_03, COD_IDEN_SEGUR_04, "
    lSelect = lSelect & " COD_IDEN_SEGUR_05, COD_IDEN_SEGUR_06, COD_IDEN_SEGUR_07, "
    lSelect = lSelect & " COD_IDEN_SEGUR_08, COD_IDEN_SEGUR_09, COD_IDEN_SEGUR_10, "
    lSelect = lSelect & " PER_TOT_SEGUR, COD_IDEN_PROTEC_01, COD_IDEN_PROTEC_02, "
    lSelect = lSelect & " COD_IDEN_PROTEC_03, COD_IDEN_PROTEC_04, COD_IDEN_PROTEC_05, "
    lSelect = lSelect & " COD_IDEN_PROTEC_06, COD_IDEN_PROTEC_07, COD_IDEN_PROTEC_08, "
    lSelect = lSelect & " COD_IDEN_PROTEC_09, COD_IDEN_PROTEC_10, PER_TOT_PROTEC, "
    lSelect = lSelect & " " & lNomeTab & "_LOC.COD_PRODUTO, COD_REGIAO, COD_IDEN_TARIF, "
    lSelect = lSelect & " " & lNomeTab & "_LOC.NUM_ENDO AS NUM_ENDO2, "
    lSelect = lSelect & " TALC.Ind_VerbaSep , TALC.Val_Vr_Lc, TALC.Val_LMG_Lc, TALC.Per_IS_Vr_Lc , TALC.Coef_IS_Vr_Lc, "
    lSelect = lSelect & " Tip_Plano_Vida,"
    lSelect = lSelect & " Cod_Plano_Cobertura,"
    lSelect = lSelect & " Qtd_Vidas,"
    lSelect = lSelect & " Val_LMG_Vida,"
    lSelect = lSelect & " Ind_Aux_Funeral,"
    lSelect = lSelect & " Val_LMG_Funeral,"
    lSelect = lSelect & " Ind_Inc_Conjuge,"
    lSelect = lSelect & " Val_Aux_Cesta, "
    lSelect = lSelect & " TIP_COBERT, "
    lSelect = lSelect & " POSSUI_ELEVADOR, "
    lSelect = lSelect & " Ind_VerbaSep, "
    lSelect = lSelect & " (SELECT top 1 Per_Desc_LMI FROM RamosDiversos.DBO.TAB_PED A WITH (NOLOCK) WHERE " & lNomeTab & "_LOC.NUM_APOL = A.NUM_APOL AND " & lNomeTab & "_LOC.NUM_ENDO = A.NUM_ENDO) AS Per_Desc_LMI "
    lSelect = lSelect & " FROM " & lNomeTab & "_LOC WITH (NOLOCK) INNER JOIN " & lNomeTab & " WITH (NOLOCK) ON "
    lSelect = lSelect & " " & lNomeTab & "_LOC.NUM_APOL = " & lNomeTab & ".NUM_APOL "
    If (M46V999.gPrograma = "D046O003") Then
        lSelect = lSelect & " INNER JOIN " & lNomeTab & "_CORR ON "
        lSelect = lSelect & " " & lNomeTab & "_CORR.NUM_APOL = " & lNomeTab & ".NUM_APOL AND "
        lSelect = lSelect & " " & lNomeTab & "_CORR.NUM_ENDO = " & lNomeTab & ".NUM_ENDO AND "
        lSelect = lSelect & " Cod_corr in (SELECT * FROM [Ged].[dbo].[Func_BuscaCorrGrupoAtivo] ('" & Format(M46V999.gCod_Corretor, "0000000") & "', 112))"
    End If

    'lSelect = lSelect & " LEFT JOIN     GED.dbo.Tab_Apol_Loc_Compl"
    lSelect = lSelect & " LEFT JOIN " & lNomeTab & "_Loc_Compl AS TALC"
    lSelect = lSelect & "   ON          TALC.NUM_PI    = " & lNomeTab & "_Loc.NUM_PI"
    lSelect = lSelect & "       AND     TALC.NUM_ITEM  = " & lNomeTab & "_Loc.NUM_ITEM"

    lSelect = lSelect & " Where " & lNomeTab & "_LOC.Num_Apol = " & pNumApolice

    If M46V999.gTipUsuario = e_Tipo_Usu_Corretor Then
        lSelect = lSelect & " AND " & lNomeTab & "_LOC.NOM_LOGR_RISCO not like '%CONFORME ESPEC%'"
    End If


    lSelect = lSelect & " AND [COD_DIZER] NOT IN (4,10)          "
    If pCod_Dizer = e_CodDizerAlteracao Or pCod_Dizer = e_CodDizerexclusao Then
        lSelect = lSelect & " AND " & lNomeTab & "_LOC.NUM_ITEM = " & pIteApolice
    End If
    lSelect = lSelect & " ORDER BY " & lNomeTab & "_LOC.NUM_ITEM, " & lNomeTab & ".DAT_EMIS desc"


    'b) Abre arquivo.
    If M46V999.gfAbrBasDados("P00APOL", M46V999.gTipP00APOL, M46V999.gEndP00APOL, bdP00APOL, pMensagem) = False Then
        Exit Function
    End If
    'c) Obt�m registro.
    If M46V999.gfObtRegistro(bdP00APOL, lSelect, rsP00APOL, pMensagem) = False Then
        Call M46V999.gpFecha2(bdP00APOL)
        Exit Function
    End If

    'd) N�o encontrou registro de ap�lice.
    If rsP00APOL.EOF = True Then

                '*Luiz Alberto em 15/01/2016
                lSelect = ""
                lSelect = "SELECT [TPD].[NUM_APOL] "
                lSelect = lSelect & " FROM [GED].[DBO].[TAB_PED]  AS [TPD] WITH(NOLOCK) "
                lSelect = lSelect & " INNER JOIN [GED]..[TAB_CTRL_EMIS] AS [TCE] WITH (NOLOCK) "
                lSelect = lSelect & " ON [TPD].[NUM_APOL]   = [TCE].[NUM_APOL] "
                lSelect = lSelect & " AND [TPD].[NUM_ENDO]  = [TCE].[NUM_ENDO] "
                lSelect = lSelect & " AND [TPD].[COD_RAMO]  = [TCE].[COD_RAMO] "
                lSelect = lSelect & " Where [TPD].[Num_Apol] <> 0 "
                lSelect = lSelect & " AND [TPD].[COD_RAMO]  IN (112,113,114) "
                lSelect = lSelect & " AND [TCE].[COD_IDEN_STAT_ACAO] = 70 "
                lSelect = lSelect & " AND [TPD].[NUM_APOL]  = " & pNumApolice
                'e)Obt�m registro.
                If M46V999.gfObtRegistro(bdP00APOL, lSelect, rsP00APOL, pMensagem) = False Then
                    Call M46V999.gpFecha2(bdP00APOL)
                    Exit Function
                End If
        If rsP00APOL.EOF = True Then
            '===1) Se n�o encontrou registro, verifica se tem ap�lice.
            If M46V999.gfPreenchido(pIteApolice) Then
                lSelect = "SELECT * FROM " & lNomeTab & "_LOC WITH (NOLOCK) "
                lSelect = lSelect & " Where Num_Apol = " & pNumApolice
                lSelect = lSelect & " order by NUM_ENDO desc "
                If M46V999.gfObtRegistro(bdP00APOL, lSelect, rsP00APOL, pMensagem) = False Then
                    Call M46V999.gpFecha2(bdP00APOL)
                    Exit Function
                End If
            End If
            
            If rsP00APOL.EOF = True Then
                                pMensagem = "FNL0012 - Ap�lice " & pNumApolice & " n�o localizada."
                Else
                    pMensagem = "FNL0013 - Item  " & pIteApolice & " n�o localizado na ap�lice " & _
                                pNumApolice & "."
                End If
        Else
            pMensagem = "FNL0012 - Ap�lice " & pNumApolice & " Cancelada - N�o dispon�vel para endosso."
        End If

        Call M46V999.gpFecha2(bdP00APOL)
        Exit Function
    End If

    If pCod_Dizer = 10 And rsP00APOL.RecordCount > 1 And lIndExclusao = True Then
        pCod_Dizer = 4
        GoTo BuscaApolice
    End If

    If rsP00APOL("dat_ter_vig") < M46V999.gHoje And M46V999.gTipUsuario = e_Tipo_Usu_Corretor Then
        pMensagem = "FNL0013 - Ap�lice vencida."
        Call M46V999.gpFecha2(bdP00APOL)
        Exit Function
    End If

    'DBIANCO - VERIFICAR COM A SIMONY
    '    If rsP00APOL("NOM_LOGR_RISCO") Like "*CONFORME ESPEC*" Then
    '        pMensagem = "FNL0013 - Ap�lice n�o dispon�vel para realiza��o de endosso. Solicitar endosso atrav�s do YasudaNet."
    '        Call M46V999.gpFecha2(bdP00APOL)
    '        Exit Function
    '    End If
    '3. Posiciona nosso n�mero.
    If M46V999.gfPreenchido(pNosNumero) Then
        lNosNumero = pNosNumero
    Else
        If M46V999.gfCalNosNumero(lNosNumero, pMensagem) = False Then
            Call M46V999.gpFecha2(bdP00APOL)
            Exit Function
        End If
        pNosNumero = lNosNumero
    End If

    '1. Abre base de dados.
    '    If M46V999.gfAbrBasDados("P00MPED", M46V999.gTipP00MPED, M46V999.gEndP00MPED, bdP00MPED, pMensagem) = False Then
    '        Call M46V999.gpGraLog(2, pMensagem)
    '        Exit Function
    '        'End
    '    End If

    '4. Posiciona n�mero do pedido.
    'If M46V999.gfCalNumPedido(lNumPedido, pMensagem, bdP00APOL) = False Then 'Alterado porque nao existe a tab_ped_ult no bdP00APOL, verificar
    'Incluido valida��o para n�o gerar novo numero de pedido quando da inclus�o de itens atraves dos endossos (12 - altera��o e 4 - Exclus�o)
    If pStcPed Is Nothing Then
        If M46V999.gfCalNumPedido(lNumPedido, pMensagem, bdP00MPED) = False Then
            Call M46V999.gpFecha2(bdP00APOL)
            Exit Function
        End If
    Else
        lNumPedido = pStcPed.Num_Ped
    End If

    M46V999.gPrefixo = "Cota��o " & lNumPedido & ". "

    '5. Posiciona vig�ncia.
    rsP00APOL.MoveFirst

    lVigDatIni = M46V999.gHoje
    lVigDatTer = rsP00APOL!Dat_Ter_Vig

    'Posicionar o mesmo plano
    M46V111.gCod_plano = rsP00APOL!Cod_Estip

    If M46V111.gCod_plano <> 10 And M46V111.gCod_plano <> 11 And rsP00APOL!Cod_Iden_Tarif < M46V999.TarifaPrecificacaoMaritima Then
        lCodProduto = 1076
    
    Else
        If M46V111.gCod_plano > 1000 Then
            lCodProduto = rsP00APOL!Cod_Iden_Tarif
        End If
    End If
    
    '6. Posiciona produto que atenda � vig�ncia.
    If M46V999.gfBusProduto(lVigDatIni, lCodProduto, lCodProtocolo, lCodSequencia, lDescTarifa, pMensagem, rsP00APOL("Cod_Ramo")) = _
       False Then
        Call M46V999.gpFecha2(bdP00APOL)
        Exit Function
    End If
    
    If lCodProduto = 0 Then
        pMensagem = "FGE0088 - N�o h� produto (tarifa) dispon�vel para c�lculo da renova��o da " & _
                    "ap�lice " & pNumApolice & ", de acordo com a data de in�cio de vig�ncia (" & _
                    M46V999.gfForData(2, lVigDatIni) & ")." & vbLf & vbLf & "Aguardar novo produto (tarifa)." & _
                    vbLf & vbLf & "Caso seja necess�rio, o c�lculo deve ser feito pela Yasuda."
        Exit Function
    End If
'    If rsP00APOL!Cod_Estip > 2000 And rsP00APOL!Cod_Estip < 5000 And (lCodProduto Mod 2) = 1 Then
'        lCodProduto = lCodProduto - 1
'    End If

    M46V999.gCodProduto = lCodProduto

    If pCadastraComum = True Then
        '7. Grava registro de renova��o Yasuda.
        If mfGraRenComum(rsP00APOL, lNosNumero, CLng(lNumPedido), lVigDatIni, lVigDatTer, lCodProduto, _
                         CInt(lCodProtocolo), CInt(lCodSequencia), pCod_Dizer, bdP00APOL, "", pNumApolice, pMensagem, lStcPed, tagDynaTrace) = False Then
            Call M46V999.gpFecha2(bdP00APOL)
            Exit Function
        End If
    End If

    If lStcPed Is Nothing Then
        Set lStcPed = pStcPed
    End If
    '8. Grava registros de itens, coberturas e descontos.
    'Caso seja encosso de inclus�o de local de risco n�o inclui item
    If pCod_Dizer <> 1 Then
        rsP00APOL.MoveLast      'Posiciona quantidade de registros.
        rsP00APOL.MoveFirst
        lNumEndo = 0
        While Not rsP00APOL.EOF
            If Not mfGraRenItem(rsP00APOL, lNosNumero, bdP00APOL, _
                                pMensagem, CLng(lNumPedido), pCod_Dizer, "", pNumApolice, lStcPed, tagDynaTrace) Then
                Call M46V999.gpGraLog(0, pMensagem)
                Exit Function
            End If
            rsP00APOL.MoveNext
        Wend
    Else
        'Acha �ltimo Item
        rsP00APOL.MoveLast      'Posiciona quantidade de registros.
        rsP00APOL.MoveFirst
        lNumItem = 0
        While Not rsP00APOL.EOF
            If rsP00APOL("Num_Item") > lNumItem Then
                lNumItem = rsP00APOL("Num_Item")
            End If
            rsP00APOL.MoveNext
        Wend
        lNumItem = lNumItem + 1
        'Abre base de dados.
        If M46V999.gfAbrBasDados("P00MPED", M46V999.gTipP00MPED, M46V999.gEndP00MPED, bdP00MPED, pMensagem) = False Then
            Exit Function
        End If
        'Grava novo registro de item.
        If M46V999.gfGraItem(bdP00MPED, lNosNumero, lNumPedido, lNumItem, pMensagem, 300, , , pNumApolice) = False Then
            Exit Function
        End If
        Call M46V999.gpFecha2(bdP00MPED)
    End If
    Call M46V999.gpFecha2(bdP00APOL)

    '9. Grava registro de comunica��o.
    If M46V111.gfGraRegComunicacao(lNosNumero, pMensagem) = False Then
        Call M46V999.gpGraLog(1, pMensagem)
        Exit Function
    End If

    Set bdP00APOL = Nothing
    Set bdP00MPED = Nothing
    Set rsP00APOL = Nothing
    Set rsRet = Nothing
    Set lStcPed = Nothing

    gfGraEndossoYasuda = True
    Exit Function
ControleErro:
    Set bdP00APOL = Nothing
    Set bdP00MPED = Nothing
    Set rsP00APOL = Nothing
    Set rsRet = Nothing
    Set lStcPed = Nothing
    Call M46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call M46V999.gpGraLog(0, "gfGraRenYasuda ERRO - " & Err & " " & Err.Description & " " & pMensagem)
End Function


Private Function mfGraRenComum(pRstApol As clsYasRecordSet, pNosNumero As String, pNumPedido As Long, _
                               pVigDatIni As Long, pVigDatTer As Long, pCodProduto As Integer, _
                               pCodProtocolo As Integer, pCodSequencia As Integer, pCod_Dizer As Integer, _
                               pBdP00APOL As ADODB.Connection, pNumCotac_LMI As String, ByVal pNumApolice As String, _
                               pMensagem As String, pStcPed As stcA46V702B, tagDynaTrace As String) As Boolean
'Fun��o: Grava os registros de comum com os dados da ap�lice (Renova��o Yasuda).
'Par�metros de entrada :pNum_Apol........N�mero da ap�lice que est� sendo renovada.
'                       pNosNumero.......Nosso n�mero.
'                       pNumPedido.......N�mero do pedido.
'                       pVigDatIni.......Data de in�cio da vig�ncia.
'                       pVigDatTer.......Data de t�rmino da vig�ncia.
'                       pCodProduto......C�digo do produto.
'                       pCodProtocolo....C�digo do protocolo.
'                       pCodSequencia....C�digo da sequencia.
'                       pBdP00APOL.......Arquivo de conex�o
'Par�metro de sa�da:    pMensagem........Mensagem de erro.

    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "mfGraRenComum"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Dim lSql As String        'Consulta SQL.
    Dim rsAPOL As clsYasRecordSet    'Registro.
    Dim lobjPedido As clsA46V702A         'Objeto da classe clsA46V702A (TAB_PED).
    Dim lobjstcPedido As stcA46V702B            'Objeto da classe de estrutura stcA46V702B (TAB_PED).
    Dim lobjCorr As clsA46V705A       'Objeto da classe clsA46V705A (TAB_PED_CORR).
    Dim lObjstcCorr As stcA46V705B          'Objeto da classe de estrutura stcA46V705B (TAB_PED_CORR).
    Dim lobjCoss As clsA46V706A       'Objeto da classe clsA46V705A (TAB_PED_CORR).
    Dim lobjstcCoss As stcA46V706B          'Objeto da classe de estrutura stcA46V705B (TAB_PED_CORR).
    Dim lObjEnd As clsA46V707A      'Objeto da classe clsA46V707A (TAB_PED_ENDER).
    Dim lObjstcEnd As stcA46V707B         'Objeto da classe de estrutura stcA46V707B (TAB_PED_ENDER).
    Dim lobjCtrlEmis As clsA46V712A           'Objeto da classe clsA46V712A (TAB_CTRL_EMIS).
    Dim lobjstcCtrlEmis As stcA46V712B              'Objeto da classe de estrutura stcA46V712B (TAB_CTRL_EMIS).
    Dim lobjJuros As clsA46V711A        'Objeto da classe clsA46V711A (TAB_PED_JUROS).
    Dim lobjAuto As clsA46V719A       'Objeto da classe clsA46V707A (TAB_PED_AUTO).
    Dim lobjEspec As clsA46V726A        'Objeto da classe clsA46V721A (TAB_PED_ESPEC).
    Dim lobjstcEspec As stcA46V726B           'Objeto da classe de estrutura stcA46V721B (TAB_PED_ESPEC).
    Dim clsA99V001 As clsA99V001
    Dim lNomeTab As String
    Dim lDigito As String
    Dim lAuxiliar As String     'Auxiliar.
    Dim lPropTabela As String
    Dim lCodOrigem As Integer   ' 1 - yasusa, 2 - maritima
    Dim lPos As Integer

    On Error GoTo ControleErro
    Call M46V999.gfObtSYASMULTINI(lPropTabela)

    If Not pRstApol.EOF Then
        Set lobjPedido = New clsA46V702A
        Set lobjPedido.M46V999 = M46V999

        Set lobjstcPedido = New stcA46V702B
        Set lobjstcPedido.M46V999 = M46V999

        With lobjstcPedido
            .nosso_numero = pNosNumero
            .Num_Ped = pNumPedido
            .Cod_Ramo = M46V999.gfTratarNulo(pRstApol("Cod_Ramo"), tpoNumerico)

            M46V999.gCod_Ramo = .Cod_Ramo

            lCodOrigem = M46V999.gfTratarNulo(pRstApol("Cod_Org"), tpoNumerico)
            If lCodOrigem = 2 Then
                .Cod_Org = 0
                .Num_Apol_An = 0
            Else
                .Cod_Org = M46V999.gfTratarNulo(pRstApol("Cod_Org"), tpoNumerico)
                .Num_Apol_An = pRstApol("Num_Apol")
            End If
            .Cod_Unid_Prod = M46V999.gfTratarNulo(pRstApol("Cod_Unid_Prod"), tpoNumerico)

            'ficha - 801312 (ajustado em 23/04)
            '.Cod_Prod = M46V999.gfTratarNulo(pRstApol("Cod_Prod"), tpoNumerico)
            .Cod_Prod = 0
            .Cod_Segurado = M46V999.gfTratarNulo(pRstApol("Cod_Segurado"), tpoNumerico)
            .Nom_Segurado = M46V999.gfTratarNulo(pRstApol("Nom_Segurado"), tpoString)
            .Dsc_Compl_Segurado = M46V999.gfTratarNulo(pRstApol("Dsc_Compl_Segurado"), tpoString)
            .Tip_Pessoa = M46V999.gfTratarNulo(pRstApol("Tip_Pessoa"), tpoNumerico)
            .Num_Cgc_Cpf = M46V999.gfTratarNulo(pRstApol("Num_Cgc_Cpf"), tpoNumerico)
            .Cod_Tip_Logr = M46V999.gfTratarNulo(pRstApol("Cod_Tip_Logr"), tpoString)
            .Nom_Logr = M46V999.gfTratarNulo(pRstApol("Nom_Logr"), tpoString)
            .Dsc_Compl_Logr = M46V999.gfTratarNulo(pRstApol("Dsc_Compl_Logr"), tpoString)
            .Nom_Bairro = M46V999.gfTratarNulo(pRstApol("Nom_Bairro"), tpoString)
            .Nom_Cid = M46V999.gfTratarNulo(pRstApol("Nom_Cid"), tpoString)
            .Sig_UF = M46V999.gfTratarNulo(pRstApol("Sig_UF"), tpoString)
            .Num_CEP = M46V999.gfTratarNulo(pRstApol("Num_Cep"), tpoNumerico)
            .Num_Fone = M46V999.gfTratarNulo(pRstApol("Num_Fone"), tpoNumerico)
            .Per_Desc_LMI = M46V999.gfTratarNulo(pRstApol("Per_Desc_LMI"), tpoNumerico)
            .Dat_Ini_Vig = pVigDatIni
            .Dat_Ter_Vig = pVigDatTer
            .Cod_Iden_Praz = e_TipoPrazo_ProRata
            .Cod_User_Liber = M46V999.gNomUsuario6
            .Dat_Liber = 0
            .Hor_Liber = 0
            .Cod_Sit_Ped = e_SitPedido_SemCalculo
            .Cod_Tip_Doc = e_TipDoc_Cotacao
            'If pCod_Dizer = 0 Then
            .Cod_Produto = pCodProduto
            'Else
            '    .Cod_Produto = M46V999.gfTratarNulo(pRstApol("COD_IDEN_TARIF"), tpoNumerico)
            'End If
            .Cod_Protocolo = pCodProtocolo
            .Cod_Sequencia = pCodSequencia
            If pCod_Dizer = 0 Then .Cod_Reg = 1
            If pCod_Dizer = 1 Then .Cod_Reg = 101
            If pCod_Dizer = 2 Then .Cod_Reg = 101
            If pCod_Dizer = 3 Then .Cod_Reg = 101
            If pCod_Dizer = 4 Then .Cod_Reg = 106
            If pCod_Dizer = 5 Then .Cod_Reg = 101
            If pCod_Dizer = 6 Then .Cod_Reg = 120
            If pCod_Dizer = 7 Then .Cod_Reg = 101
            If pCod_Dizer = 10 Then .Cod_Reg = 107
            If pCod_Dizer = 11 Then pCod_Dizer = 18
            If pCod_Dizer = 12 Then .Cod_Reg = 101
            If pCod_Dizer = 18 Then .Cod_Reg = 120
            .Cod_Dizer = pCod_Dizer

            If .Cod_Produto >= 1080 And pRstApol("COD_ESTIP") < 2000 Then
                'CR93893N�o faz essa regra para endosso nem para renova��o de apolice emitida na 1080 e apos.
                If pCod_Dizer = 0 And pRstApol("COD_IDEN_TARIF") < 1080 Then
                    If pRstApol("VAL_VR") > 2000000 And .Cod_Ramo = 112 Then 'SETAR EMPRESARIAL SUPER QUANDO MAIOR QUE 2 MILHOES
                        .Cod_Estip = M46V999.gfPlanoBasico(.Cod_Ramo, 2)
                    Else
                        .Cod_Estip = M46V999.gfPlanoBasico(.Cod_Ramo, .Cod_Estip)
                    End If
                Else
                    .Cod_Estip = pRstApol("COD_ESTIP")
                End If
            Else
                If .Cod_Produto >= 1080 And Not .getPlano.gVerificaPlanoAtivo And .Cod_Dizer = 0 Then
                    M46V111.gCod_plano = 10
                    .Cod_Estip = M46V111.gCod_plano
                Else
                    .Cod_Estip = M46V999.gfTratarNulo(pRstApol("COD_ESTIP"), tpoNumerico)
                End If
            End If

            If M46V111.gNum_Pi = 0 Then
                .NUM_PI = Right(Year(Date), 2) & Format(pNumPedido, "00000000")
            Else
                .NUM_PI = M46V111.gNum_Pi
                M46V111.gNum_Pi = .NUM_PI
            End If

            If pCod_Dizer = 0 Then
                .Val_Cust_Emis = M46V999.gProCusApolice
            Else
                .Val_Cust_Emis = M46V999.gProCusEndosso
            End If
            .Tip_Cust_Emis = 1
            If pCod_Dizer <> 0 Then
                .Val_Parc_Min = M46V999.gProPreMinParcelaEndosso
            Else
                .Val_Parc_Min = M46V999.gProPreMinParcela
            End If
            .Cod_Agen_Bas_Cor = M46V999.gfTratarNulo(pRstApol("Cod_Agen_Bas_Cor"), tpoNumerico)
            If .Cod_Agen_Bas_Cor <> 0 Then
                .Cod_Bco_Cobr = M46V999.gfTratarNulo(pRstApol("Cod_Bco_Cobr"), tpoNumerico)
            End If
            If M46V999.gTipUsuario = e_Tipo_Usu_Corretor Or M46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda Then
                lDigito = 0
                Set clsA99V001 = New clsA99V001
                Set clsA99V001.M46V999 = M46V999

                Select Case .Cod_Bco_Cobr
                Case 1  'Projeto 11080 - Debito automatico - Banco do Brasil
                    Call clsA99V001.DIG_BANCOBRASIL(Format(M46V999.gfTratarNulo(pRstApol("Num_Cta_Bas_Cor"), tpoNumerico), "00000000"), lDigito, lAuxiliar)
                Case 237
                    Call clsA99V001.DIG_BRADESCO(Format(M46V999.gfTratarNulo(pRstApol("Num_Cta_Bas_Cor"), tpoNumerico), "0000000"), lDigito, lAuxiliar)
                Case 341
                    Call clsA99V001.DIG_ITAU(Format(.gAgenciaDig, "0000"), Format(M46V999.gfTratarNulo(pRstApol("Num_Cta_Bas_Cor"), tpoNumerico), "00000"), lDigito, lAuxiliar)
                Case 347
                    If .gAgenciaDig >= 1000 Then
                        Call clsA99V001.DIG_SUDAMERIS_NOVO(Format(.gAgenciaDig, "0000"), Format(M46V999.gfTratarNulo(pRstApol("Num_Cta_Bas_Cor"), tpoNumerico), "0000000"), lDigito, lAuxiliar)
                    Else
                        Call clsA99V001.DIG_SUDAMERIS(Format(.gAgenciaDig, "000"), Format(M46V999.gfTratarNulo(pRstApol("Num_Cta_Bas_Cor"), tpoNumerico), "000000000"), lDigito, lAuxiliar)
                    End If
                Case 409
                    Call clsA99V001.DIG_UNIBANCO(Format(.gAgenciaDig, "0000"), Format(M46V999.gfTratarNulo(pRstApol("Num_Cta_Bas_Cor"), tpoNumerico), "000000"), lDigito, lAuxiliar)
                Case 479
                    Call clsA99V001.DIG_BOSTON(Format(M46V999.gfTratarNulo(pRstApol("Num_Cta_Bas_Cor"), tpoNumerico), "0000000"), lDigito, lAuxiliar)
                Case 356
                    Call clsA99V001.DIG_REAL(Format(.gAgenciaDig, "0000"), Format(M46V999.gfTratarNulo(pRstApol("Num_Cta_Bas_Cor"), tpoNumerico), "0000000"), lDigito, lAuxiliar)
                End Select
                Set clsA99V001 = Nothing
                If Not IsNumeric(lDigito) Then lDigito = 0
                .Num_Cta_Bas_Cor = M46V999.gfTratarNulo(pRstApol("Num_Cta_Bas_Cor"), tpoNumerico) & lDigito
            Else
                .Num_Cta_Bas_Cor = M46V999.gfTratarNulo(pRstApol("Num_Cta_Bas_Cor"), tpoNumerico)
            End If

            If .Cod_Bco_Cobr = 409 Then
                .Cod_Bco_Cobr = 341
                .Cod_Agen_Bas_Cor = 0
                .Num_Cta_Bas_Cor = 0
            End If

            .LMI_Unico = M46V999.gfTratarNulo(pRstApol("LMI_Unico"), tpoNumerico)
            'FICHA 900579 - CARREGAR CAMPOS DO SISCAD
            '------------------------------------------------------------------------------------------------
            If M46V999.gTipUsuario = e_Tipo_Usu_Corretor Or M46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda Then
                On Error Resume Next
                .Tip_Pessoa = M46V999.gfTratarNulo(pRstApol("Tip_Pessoa"), tpoNumerico)
                .Cod_Ramo_Ativ = M46V999.gfTratarNulo(pRstApol("Cod_Ramo_Ativ"), tpoNumerico)
                .Cod_Doc = M46V999.gfTratarNulo(pRstApol("Cod_Doc"), tpoNumerico)
                .Num_Doc = M46V999.gfTratarNulo(pRstApol("Num_Doc"), tpoString)
                .Org_Exp = M46V999.gfTratarNulo(pRstApol("Org_Exp"), tpoString)
                .Dat_Emis_Doc = M46V999.gfTratarNulo(pRstApol("Dat_Emis_Doc"), tpoNumerico)
                'Ficha 900417
                If M46V999.gCodProduto >= 968 Then
                    .Nom_Empresa = M46V999.gfTratarNulo(pRstApol("Nom_Profissao"), tpoString)
                Else
                    .Nom_Empresa = M46V999.gfTratarNulo(pRstApol("Nom_Empresa"), tpoString)
                End If
                .Sexo = M46V999.gfTratarNulo(pRstApol("Sexo"), tpoString)
                Err.Clear
                On Error GoTo ControleErro
            End If
            '------------------------------------------------------------------------------------------------

            'Ficha 1000496
            'Corre��o de Produ��o - Unidade
            If M46V999.gBaseLocal = False Then
                lNomeTab = "GED.DBO.TAB_CORRECAO_PROD"
                lSql = " Select top 1 cod_unid_acert from " & lNomeTab & " WITH (NOLOCK) Where Num_Apol = " & pRstApol("Num_Apol")
                lSql = lSql & " and cod_unid_acert > 0 and niv_corr = 1  "
                lSql = lSql & " AND DAT_TER  = 99999999"

                If Not M46V999.gfObtRegistro(pBdP00APOL, lSql, rsAPOL, pMensagem) Then
                    GoTo SAIDA
                End If

                If Not rsAPOL.EOF Then
                    .Cod_Unid_Prod = M46V999.gfTratarNulo(rsAPOL("cod_unid_acert"), tpoNumerico)
                End If

                'Corre��o de Produ��o - Produtor
                lNomeTab = "GED.DBO.TAB_CORRECAO_PROD"
                lSql = " Select top 1 cod_prod_acert from " & lNomeTab & " WITH (NOLOCK) Where Num_Apol = " & pRstApol("Num_Apol")
                lSql = lSql & " and cod_prod_acert > 0 and niv_corr = 3  "
                lSql = lSql & " AND DAT_TER  = 99999999"

                If Not M46V999.gfObtRegistro(pBdP00APOL, lSql, rsAPOL, pMensagem) Then
                    GoTo SAIDA
                End If

                If Not rsAPOL.EOF Then
                    .Cod_Prod = M46V999.gfTratarNulo(rsAPOL("cod_prod_acert"), tpoNumerico)
                End If
            End If

            If pNumCotac_LMI <> "" Then    '1100177 - Renova��o Apolice LMI
                .LMI_Unico = 1
                .Num_Apol_An = pNumApolice
            End If


            If pCod_Dizer = 0 Then
                .Qtd_Parc = 11
            ElseIf pCod_Dizer = 4 Or pCod_Dizer = 10 Then
                .Qtd_Parc = 0
            Else
                .Qtd_Parc = 1
            End If
        End With

        If Not lobjPedido.mfIncluir(lobjstcPedido, lPropTabela, "", tagDynaTrace) Then
            pMensagem = lobjPedido.UltimoErro
            GoTo SAIDA
        End If

        If M46V999.gBaseLocal = True Then
            lNomeTab = "TAB_APOL"
        Else
            If pNumCotac_LMI <> "" Then    '1100177 - Renova��o Apolice LMI
                lNomeTab = "RAMOSDIVERSOS.DBO.TAB_PED"
            Else
                'lNomeTab = "GED.DBO.TAB_PED"
                lNomeTab = "TAB_APOL"
            End If
        End If

        'Corretores
        lSql = " Select * from " & lNomeTab & "_CORR WITH (NOLOCK) Where "
        If pNumCotac_LMI = "" Then    '1100177 - Renova��o Apolice LMI
            lSql = lSql & " Num_Apol = " & pRstApol("Num_Apol")
            lSql = lSql & " and Num_Endo = " & pRstApol("Num_Endo")
        Else
            lSql = lSql & " NOSSO_NUMERO = " & M46V999.gfFormatarCamposSQL(pNumCotac_LMI)
        End If

        If Not M46V999.gfObtRegistro(pBdP00APOL, lSql, rsAPOL, pMensagem) Then
            GoTo SAIDA
        End If

        Set lobjCorr = New clsA46V705A
        Set lobjCorr.M46V999 = M46V999

        Set lobjCorr.ConexaoBD = lobjPedido.ConexaoBD
        Do While Not rsAPOL.EOF
            Set lObjstcCorr = New stcA46V705B
            Set lObjstcCorr.M46V999 = M46V999

            With lObjstcCorr
                .Cod_Ramo = M46V999.gfTratarNulo(rsAPOL("Cod_Ramo"), tpoNumerico)
                .nosso_numero = pNosNumero
                .Num_Ped = pNumPedido
                If M46V111.gNum_Pi = 0 Then
                    .NUM_PI = Right(Year(Date), 2) & Format(pNumPedido, "00000000")
                Else
                    .NUM_PI = M46V111.gNum_Pi
                    M46V111.gNum_Pi = .NUM_PI
                End If
                .Cod_Corr = M46V999.gfTratarNulo(rsAPOL("Cod_Corr"), tpoNumerico)
                M46V999.gCod_Corretor = .Cod_Corr    '1001135 - Corretor Volks - Franquia Especial Concessionaria
                .Num_Ordem_Corr = M46V999.gfTratarNulo(rsAPOL("Num_Ordem_Corr"), tpoNumerico)

                .Per_Corr = M46V999.gfTratarNulo(rsAPOL("Per_Corr"), TpoValor)



                If .Num_Ordem_Corr = 1 Then
                    'Posiciona o Produtor vigente
                    lobjstcPedido.Cod_Prod = gfRetornaCodProdutor(.Cod_Corr, lobjstcPedido.Cod_Unid_Prod)
                    Call lobjPedido.mfAlterar(lobjstcPedido, lPropTabela, "", , lobjstcPedido.Cod_User_Liber)
                End If
            End With

            Call lobjCorr.mfIncluir(lObjstcCorr, lPropTabela, "", tagDynaTrace)
            rsAPOL.MoveNext
        Loop

        'Cosseguro
        lSql = " Select * from " & lNomeTab & "_COSS WITH (NOLOCK) Where "
        If pNumCotac_LMI = "" Then    '1100177 - Renova��o Apolice LMI
            lSql = lSql & " Num_Apol = " & pRstApol("Num_Apol")
            lSql = lSql & " and Num_Endo = " & pRstApol("Num_Endo")
        Else
            lSql = lSql & " NOSSO_NUMERO = " & M46V999.gfFormatarCamposSQL(pNumCotac_LMI)
        End If

        If Not M46V999.gfObtRegistro(pBdP00APOL, lSql, rsAPOL, pMensagem) Then
            GoTo SAIDA
        End If

        Set lobjCoss = New clsA46V706A
        Set lobjCoss.M46V999 = M46V999

        Set lobjCoss.ConexaoBD = lobjPedido.ConexaoBD
        Do While Not rsAPOL.EOF
            Set lobjstcCoss = New stcA46V706B
            Set lobjstcCoss.M46V999 = M46V999

            With lobjstcCoss
                .Cod_Ramo = M46V999.gfTratarNulo(rsAPOL("Cod_Ramo"), tpoNumerico)
                .nosso_numero = pNosNumero
                .Num_Ped = pNumPedido
                If M46V111.gNum_Pi = 0 Then
                    .NUM_PI = Right(Year(Date), 2) & Format(pNumPedido, "00000000")
                    M46V111.gNum_Pi = .NUM_PI
                Else
                    .NUM_PI = M46V111.gNum_Pi
                End If
                .Cod_Coss = M46V999.gfTratarNulo(rsAPOL("Cod_Coss"), tpoNumerico)
                .Per_Coss = M46V999.gfTratarNulo(rsAPOL("Per_Coss"), tpoNumerico)
            End With
            Call lobjCoss.mfIncluir(lobjstcCoss, lPropTabela, "", tagDynaTrace)
            rsAPOL.MoveNext
        Loop

        'Endere�o
        lSql = " Select * from " & lNomeTab & "_ENDER WITH (NOLOCK) Where "
        If pNumCotac_LMI = "" Then    '1100177 - Renova��o Apolice LMI
            lSql = lSql & " Num_Apol = " & pRstApol("Num_Apol")
            lSql = lSql & " and Num_Endo = " & pRstApol("Num_Endo")
        Else
            lSql = lSql & " NOSSO_NUMERO = " & M46V999.gfFormatarCamposSQL(pNumCotac_LMI)
        End If
        If Not M46V999.gfObtRegistro(pBdP00APOL, lSql, rsAPOL, pMensagem) Then
            GoTo SAIDA
        End If

        Set lObjEnd = New clsA46V707A
        Set lObjEnd.M46V999 = M46V999

        Set lObjEnd.ConexaoBD = lobjPedido.ConexaoBD
        If Not rsAPOL.EOF Then
            Set lObjstcEnd = New stcA46V707B
            Set lObjstcEnd.M46V999 = M46V999
            With lObjstcEnd
                .nosso_numero = pNosNumero
                .Cod_Ramo = lobjstcPedido.Cod_Ramo
                .Num_Ped = lobjstcPedido.Num_Ped
                If M46V111.gNum_Pi = 0 Then
                    .NUM_PI = Right(Year(Date), 2) & Format(pNumPedido, "00000000")
                    M46V111.gNum_Pi = .NUM_PI
                Else
                    .NUM_PI = M46V111.gNum_Pi
                End If
                .Cod_Tip_Logr = M46V999.gfTratarNulo(rsAPOL("Cod_Tip_Logr"), tpoString)
                .Nom_Logr = M46V999.gfTratarNulo(rsAPOL("Nom_Logr"), tpoString)
                .Dsc_Compl_Logr = M46V999.gfTratarNulo(rsAPOL("Dsc_Compl_Logr"), tpoString)
                .Nom_Bairro = M46V999.gfTratarNulo(rsAPOL("Nom_Bairro"), tpoString)
                .Nom_Cid = M46V999.gfTratarNulo(rsAPOL("Nom_Cid"), tpoString)
                .Sig_UF = M46V999.gfTratarNulo(rsAPOL("Sig_UF"), tpoString)
                .Num_CEP = M46V999.gfTratarNulo(rsAPOL("Num_Cep"), tpoNumerico)
                .Num_Fone = M46V999.gfTratarNulo(rsAPOL("Num_Fone"), tpoNumerico)
                'Ficha 1000307
                If pCod_Dizer <> 0 Then
                    .Cod_Reg_Cons = M46V999.gfTratarNulo(rsAPOL("Cod_Reg_Cons"), tpoNumerico)
                    .Num_Reg_Cons = M46V999.gfTratarNulo(rsAPOL("Num_Reg_Cons"), tpoString)
                End If
            End With
            If Not lObjEnd.mfIncluir(lObjstcEnd, lPropTabela, "", tagDynaTrace) Then
                pMensagem = lObjEnd.UltimoErro
                GoTo SAIDA
            End If
        Else
            Set lObjstcEnd = New stcA46V707B
            Set lObjstcEnd.M46V999 = M46V999

            With lObjstcEnd
                .nosso_numero = pNosNumero
            End With
            If Not lObjEnd.mfIncluir(lObjstcEnd, lPropTabela, "", tagDynaTrace) Then
                pMensagem = lObjEnd.UltimoErro
                GoTo SAIDA
            End If
        End If

        Set lobjstcPedido.objstcPedEndereco = lObjstcEnd

        'Juros
        'Set lobjJuros = New clsA46V711A
        'Set lobjJuros.M46V999 = M46V999
        'Set lobjJuros.ConexaoBD = lobjPedido.ConexaoBD
        'If Not lobjJuros.mfCopiarPadraoParaNovo(pNosNumero) Then
        '    pMensagem = lobjJuros.UltimoErro
        '    GoTo ControleErro
        'End If

        'Controle de emiss�o
        Set lobjCtrlEmis = New clsA46V712A
        Set lobjCtrlEmis.M46V999 = M46V999

        Set lobjCtrlEmis.ConexaoBD = lobjPedido.ConexaoBD
        Set lobjstcCtrlEmis = New stcA46V712B

        lSql = " Select * from RamosDiversos.dbo.Tab_ctrl_emis_compl WITH (NOLOCK) Where "
        lSql = lSql & " Num_Apol = " & pRstApol("Num_Apol")
        lSql = lSql & " and Num_endo  = " & pRstApol("Num_endo")

        If Not M46V999.gfObtRegistro(pBdP00APOL, lSql, rsAPOL, pMensagem) Then
            pMensagem = lobjCtrlEmis.UltimoErro
            GoTo SAIDA
        End If

        If Not rsAPOL.EOF Then
            lobjstcCtrlEmis.DAT_Nascimento = M46V999.gfTratarNulo(rsAPOL("DAT_Nascimento"), tpoNumerico)
            lobjstcCtrlEmis.Estado_Civil = M46V999.gfTratarNulo(rsAPOL("Estado_Civil"), tpoNumerico)
            lobjstcCtrlEmis.Ind_Cond_Constituido = M46V999.gfTratarNulo(rsAPOL("Ind_Cond_Constituido"), tpoNumerico)
            
            If pCod_Dizer > 0 Then
                lobjstcCtrlEmis.Suprimir_Batente = M46V999.gfTratarNulo(rsAPOL("Suprimir_Batente"), TpoBoolean)
            Else
                lobjstcCtrlEmis.Suprimir_Batente = 0
            End If
            
'            If M46V999.gfTratarNulo(pRstApol("Indice_IPC"), tpoNumerico) = 0 And lobjstcPedido.Cod_Estip < 2000 And lobjstcPedido.Cod_Dizer = 0 Then
'                lobjstcCtrlEmis.Indice_IPC = M46V999.gfObtemIPC(pMensagem)
'            Else
                lobjstcCtrlEmis.Indice_IPC = 0
'            End If
        End If
        
        lobjstcCtrlEmis.nosso_numero = pNosNumero

'        If M46V999.gCod_Ramo = e_CodRamo_Riscos_diversos Then
        lobjstcCtrlEmis.NUM_PI = M46V111.gNum_Pi
        lobjstcCtrlEmis.Cod_Ramo = lobjstcPedido.Cod_Ramo
        
        lobjstcCtrlEmis.cod_origem = lCodOrigem
        If IsNumeric(lobjstcPedido.objstcPedEndereco.Dsc_Compl_Logr) Then
            lobjstcCtrlEmis.Num_Logr_Segurado = CDbl(lobjstcPedido.objstcPedEndereco.Dsc_Compl_Logr)
        Else
            If Left(lobjstcPedido.Dsc_Compl_Logr, 2) = "N." Then
                lobjstcCtrlEmis.Num_Logr_Segurado = Mid(lobjstcPedido.objstcPedEndereco.Dsc_Compl_Logr, 3, InStr(1, lobjstcPedido.objstcPedEndereco.Dsc_Compl_Logr, "-") - 3)
                lobjstcCtrlEmis.Dsc_Compl_Segurado = Left(Trim(Mid(lobjstcPedido.objstcPedEndereco.Dsc_Compl_Logr, InStr(1, lobjstcPedido.objstcPedEndereco.Dsc_Compl_Logr, "-") + 1)), 13)
            Else
                lPos = 0
                lPos = InStr(1, lobjstcPedido.objstcPedEndereco.Dsc_Compl_Logr, " ")
                If lPos > 0 Then
                    If IsNumeric(Left(lobjstcPedido.objstcPedEndereco.Dsc_Compl_Logr, lPos)) Then
                        lobjstcCtrlEmis.Num_Logr_Segurado = Left(lobjstcPedido.objstcPedEndereco.Dsc_Compl_Logr, lPos)
                        lobjstcCtrlEmis.Dsc_Compl_Segurado = Left(Trim(Mid(lobjstcPedido.objstcPedEndereco.Dsc_Compl_Logr, lPos + 1)), 13)
                    End If
                End If
            End If
        End If
        
        If IsNumeric(lobjstcPedido.Dsc_Compl_Logr) Then
            lobjstcCtrlEmis.Num_Logr_Cobranca = CDbl(lobjstcPedido.Dsc_Compl_Logr)
        Else
            If Left(lobjstcPedido.Dsc_Compl_Logr, 2) = "N." Then
                lobjstcCtrlEmis.Num_Logr_Cobranca = Mid(lobjstcPedido.Dsc_Compl_Logr, 3, InStr(1, lobjstcPedido.Dsc_Compl_Logr, "-") - 3)
                lobjstcCtrlEmis.Dsc_Compl_Cobranca = Left(Trim(Mid(lobjstcPedido.Dsc_Compl_Logr, InStr(1, lobjstcPedido.Dsc_Compl_Logr, "-") + 1)), 13)
            Else
                lPos = 0
                lPos = InStr(1, lobjstcPedido.Dsc_Compl_Logr, " ")
                If lPos > 0 Then
                    If IsNumeric(Left(lobjstcPedido.Dsc_Compl_Logr, lPos)) Then
                                lobjstcCtrlEmis.Num_Logr_Cobranca = Left(lobjstcPedido.Dsc_Compl_Logr, lPos)
                                lobjstcCtrlEmis.Dsc_Compl_Cobranca = Left(Trim(Mid(lobjstcPedido.Dsc_Compl_Logr, lPos + 1)), 13)
                        End If
                End If
                End If
        End If
        If lCodOrigem = 2 Then 'Renova��o Maritima
            lobjstcCtrlEmis.Num_Apol_Mar = pRstApol("Num_Apol")
            'lobjstcCtrlEmis.Num_Ref_Corr = pRstApol("Num_Apol")
        Else
            lobjstcCtrlEmis.Num_Apol_Ant = lobjstcPedido.Num_Apol_An
        End If
        
'        End If

        ''2016-08-29: Ajuste SISCOTA
        If lobjstcCtrlEmis.Tip_Cotac = "" And Not UCase(App.EXEName) Like "D46O003*" Then
            lobjstcCtrlEmis.Tip_Cotac = 2
        End If

        If Not rsAPOL.EOF Then
            'Restitui��o - Projeto capita��o de dados bancarios
            lobjstcCtrlEmis.Cod_Banco_Cre = M46V999.gfTratarNulo(rsAPOL("Cod_Banco_Cre"), tpoNumerico)
            lobjstcCtrlEmis.Cod_Agenc_Cre = M46V999.gfTratarNulo(rsAPOL("Cod_Agenc_Cre"), tpoNumerico)
            lobjstcCtrlEmis.Cpf_Cnpj_Cre = M46V999.gfTratarNulo(rsAPOL("Cpf_Cnpj_Cre"), tpoString)
            lobjstcCtrlEmis.Dig_Agenc_Cre = M46V999.gfTratarNulo(rsAPOL("Dig_Agenc_Cre"), tpoNumerico)
            lobjstcCtrlEmis.Tip_Res_Cre = M46V999.gfTratarNulo(rsAPOL("Tip_Res_Cre"), tpoNumerico)
            lobjstcCtrlEmis.Cod_Conta_Cre = M46V999.gfTratarNulo(rsAPOL("Cod_Conta_Cre"), tpoNumerico)
            lobjstcCtrlEmis.Cod_DIG_Conta = M46V999.gfTratarNulo(rsAPOL("Cod_DIG_Conta"), tpoString)
            lobjstcCtrlEmis.Cod_DIG_Conta = UCase(lobjstcCtrlEmis.Cod_DIG_Conta)
        End If
        
        If Not lobjCtrlEmis.mfIncluir(lobjstcCtrlEmis, lPropTabela, "", tagDynaTrace) Then
            pMensagem = lobjCtrlEmis.UltimoErro
            GoTo SAIDA
        End If

        Set lobjstcPedido.objstcCtrlEmis = lobjstcCtrlEmis

        If pCod_Dizer <> 0 Then
            'Especifica��o
            lSql = " Select * from " & lNomeTab & "_ESPEC WITH (NOLOCK) Where "
            If pNumCotac_LMI = "" Then    '1100177 - Renova��o Apolice LMI
                lSql = lSql & " Num_Apol = " & pRstApol("Num_Apol")
                lSql = lSql & " and Num_Endo = " & pRstApol("Num_Endo")
            Else
                lSql = lSql & " NOSSO_NUMERO = " & M46V999.gfFormatarCamposSQL(pNumCotac_LMI)
            End If
            lSql = lSql & " AND COD_ESPEC <> 1"
            lSql = lSql & " AND COD_ESPEC <> 8585"
            If Not M46V999.gfObtRegistro(pBdP00APOL, lSql, rsAPOL, pMensagem) Then
                GoTo SAIDA
            End If


            Set lobjEspec = New clsA46V726A
            Set lobjEspec.M46V999 = M46V999
            Set lobjEspec.ConexaoBD = lobjPedido.ConexaoBD

            Do While Not rsAPOL.EOF
                Set lobjstcEspec = New stcA46V726B
                lobjstcEspec.nosso_numero = pNosNumero
                lobjstcEspec.Cod_Espec = M46V999.gfTratarNulo(rsAPOL("Cod_Espec"), tpoNumerico)
                lobjstcEspec.Cod_Ramo = M46V999.gfTratarNulo(rsAPOL("Cod_Ramo"), tpoNumerico)
                lobjstcEspec.Num_Ped = pNumPedido
                lobjstcEspec.Num_pag = M46V999.gfTratarNulo(rsAPOL("Num_pag"), tpoNumerico)
                lobjstcEspec.Qtd_vias = M46V999.gfTratarNulo(rsAPOL("Qtd_vias"), tpoNumerico)
                lobjstcEspec.Espec_lin1 = M46V999.gfTratarNulo(rsAPOL("Espec_lin1"), tpoString)
                lobjstcEspec.Espec_lin2 = M46V999.gfTratarNulo(rsAPOL("Espec_lin2"), tpoString)
                lobjstcEspec.Espec_lin3 = M46V999.gfTratarNulo(rsAPOL("Espec_lin3"), tpoString)
                lobjstcEspec.Espec_lin4 = M46V999.gfTratarNulo(rsAPOL("Espec_lin4"), tpoString)
                lobjstcEspec.Espec_lin5 = M46V999.gfTratarNulo(rsAPOL("Espec_lin5"), tpoString)
                lobjstcEspec.Espec_lin6 = M46V999.gfTratarNulo(rsAPOL("Espec_lin6"), tpoString)
                lobjstcEspec.Espec_lin7 = M46V999.gfTratarNulo(rsAPOL("Espec_lin7"), tpoString)
                lobjstcEspec.Espec_lin8 = M46V999.gfTratarNulo(rsAPOL("Espec_lin8"), tpoString)
                lobjstcEspec.Espec_lin9 = M46V999.gfTratarNulo(rsAPOL("Espec_lin9"), tpoString)
                lobjstcEspec.Espec_lin10 = M46V999.gfTratarNulo(rsAPOL("Espec_lin10"), tpoString)
                lobjstcEspec.Espec_lin11 = M46V999.gfTratarNulo(rsAPOL("Espec_lin11"), tpoString)
                lobjstcEspec.Espec_lin12 = M46V999.gfTratarNulo(rsAPOL("Espec_lin12"), tpoString)
                lobjstcEspec.Espec_lin13 = M46V999.gfTratarNulo(rsAPOL("Espec_lin13"), tpoString)
                lobjstcEspec.Espec_lin14 = M46V999.gfTratarNulo(rsAPOL("Espec_lin14"), tpoString)
                lobjstcEspec.Espec_lin15 = M46V999.gfTratarNulo(rsAPOL("Espec_lin15"), tpoString)
                lobjstcEspec.Espec_lin16 = M46V999.gfTratarNulo(rsAPOL("Espec_lin16"), tpoString)
                lobjstcEspec.Espec_lin17 = M46V999.gfTratarNulo(rsAPOL("Espec_lin17"), tpoString)
                lobjstcEspec.Espec_lin18 = M46V999.gfTratarNulo(rsAPOL("Espec_lin18"), tpoString)
                lobjstcEspec.Espec_lin19 = M46V999.gfTratarNulo(rsAPOL("Espec_lin19"), tpoString)
                lobjstcEspec.Espec_lin20 = M46V999.gfTratarNulo(rsAPOL("Espec_lin20"), tpoString)
                lobjstcEspec.NUM_PI = M46V111.gNum_Pi
                If Not lobjEspec.mfIncluir(lobjstcEspec, lPropTabela, "", tagDynaTrace) Then
                    pMensagem = lobjEspec.UltimoErro
                    GoTo SAIDA
                End If
                rsAPOL.MoveNext
            Loop
        End If

        'Prolabore
        If M46V999.gBaseLocal = True Then
            lNomeTab = ""
        Else
            lNomeTab = "RamosDiversos.dbo."
        End If

        lSql = " Select * from " & lNomeTab & "Tab_Canal_Emiss WITH (NOLOCK) Where "
        If pNumCotac_LMI = "" Then
            lSql = lSql & " Num_Apol = " & pRstApol("Num_Apol")
            lSql = lSql & " and Num_Compl  = " & pRstApol("Num_endo")
        Else
            lSql = lSql & " NOSSO_NUMERO = " & M46V999.gfFormatarCamposSQL(pNumCotac_LMI)
        End If

        If Not M46V999.gfObtRegistro(pBdP00APOL, lSql, rsAPOL, pMensagem) Then
            'Se n�o encontrou, n�o h� o que fazer.
            GoTo SAIDA
        End If

        If Not rsAPOL.EOF Then
            Dim lobjProAgenc As New clsA46V733A
            Dim lobjstcProAgenc As New stcA46V733B

            Set lobjProAgenc = New clsA46V733A
            Set lobjProAgenc.M46V999 = M46V999

            Set lobjProAgenc.ConexaoBD = lobjPedido.ConexaoBD
            Do While Not rsAPOL.EOF
                Set lobjstcProAgenc = New stcA46V733B
                Set lobjstcProAgenc.M46V999 = M46V999

                With lobjstcProAgenc
                    .nosso_numero = lobjstcPedido.nosso_numero
                    .NUM_PI = lobjstcPedido.NUM_PI
                    .NUM_SEQ = M46V999.gfTratarNulo(rsAPOL("Num_Seq"), tpoNumerico)
                    .Tip_Canal = M46V999.gfTratarNulo(rsAPOL("Tip_Canal"), tpoNumerico)
                    .Cod_Canal = M46V999.gfTratarNulo(rsAPOL("Cod_Canal"), tpoNumerico)
                    .Cod_Corretor = M46V999.gfTratarNulo(rsAPOL("Cod_Corretor"), tpoNumerico)
                    .Perc_Canal = M46V999.gfTratarNulo(rsAPOL("Perc_Canal"), tpoNumerico)
                    .Qtd_Parc_Pag = M46V999.gfTratarNulo(rsAPOL("Qtd_Parc_Pag"), tpoNumerico)
                    .Qtd_Parc_Apol = M46V999.gfTratarNulo(rsAPOL("Qtd_Parc_Apol"), tpoNumerico)
                    .Cod_Form_Pag_Ag = M46V999.gfTratarNulo(rsAPOL("Cod_Form_Pag_Ag"), tpoNumerico)
                End With
                Call lobjProAgenc.mfIncluir(lobjstcProAgenc, lPropTabela, "", tagDynaTrace)
                rsAPOL.MoveNext
            Loop
        End If

        'Prolabore
        If M46V999.gBaseLocal = True Then
            lNomeTab = ""
        Else
            lNomeTab = "RamosDiversos.dbo."
        End If

        lSql = " Select * from " & lNomeTab & "Tab_Canal_Emiss WITH (NOLOCK) Where "
        If pNumCotac_LMI = "" Then
            lSql = lSql & " Num_Apol = " & pRstApol("Num_Apol")
            lSql = lSql & " and Num_Compl  = " & pRstApol("Num_endo")
        Else
            lSql = lSql & " NOSSO_NUMERO = " & M46V999.gfFormatarCamposSQL(pNumCotac_LMI)
        End If

        If Not M46V999.gfObtRegistro(pBdP00APOL, lSql, rsAPOL, pMensagem) Then
            'Se n�o encontrou, n�o h� o que fazer.
            'GoTo SAIDA
        End If

        If Not rsAPOL.EOF Then

            Set lobjProAgenc = New clsA46V733A
            Set lobjProAgenc.M46V999 = M46V999

            Set lobjProAgenc.ConexaoBD = lobjPedido.ConexaoBD
            Do While Not rsAPOL.EOF
                Set lobjstcProAgenc = New stcA46V733B
                Set lobjstcProAgenc.M46V999 = M46V999

                With lobjstcProAgenc
                    .nosso_numero = lobjstcPedido.nosso_numero
                    .NUM_PI = lobjstcPedido.NUM_PI
                    .NUM_SEQ = M46V999.gfTratarNulo(rsAPOL("Num_Seq"), tpoNumerico)
                    .Tip_Canal = M46V999.gfTratarNulo(rsAPOL("Tip_Canal"), tpoNumerico)
                    .Cod_Canal = M46V999.gfTratarNulo(rsAPOL("Cod_Canal"), tpoNumerico)
                    .Cod_Corretor = M46V999.gfTratarNulo(rsAPOL("Cod_Corretor"), tpoNumerico)
                    .Perc_Canal = M46V999.gfTratarNulo(rsAPOL("Perc_Canal"), tpoNumerico)
                    .Qtd_Parc_Pag = M46V999.gfTratarNulo(rsAPOL("Qtd_Parc_Pag"), tpoNumerico)
                    .Qtd_Parc_Apol = M46V999.gfTratarNulo(rsAPOL("Qtd_Parc_Apol"), tpoNumerico)
                    .Cod_Form_Pag_Ag = M46V999.gfTratarNulo(rsAPOL("Cod_Form_Pag_Ag"), tpoNumerico)
                End With
                Call lobjProAgenc.mfIncluir(lobjstcProAgenc, lPropTabela, "", tagDynaTrace)
                rsAPOL.MoveNext
            Loop
        End If


        lSql = " Select * from " & lNomeTab & "Tab_Apol_Telefones WITH (NOLOCK) Where "
        If pNumCotac_LMI = "" Then
            lSql = lSql & " Num_Apol = " & pRstApol("Num_Apol")
            lSql = lSql & " and Num_Endo = " & pRstApol("Num_Endo")
        Else
            lSql = lSql & " NOSSO_NUMERO = " & M46V999.gfFormatarCamposSQL(pNumCotac_LMI)
        End If

        If Not M46V999.gfObtRegistro(pBdP00APOL, lSql, rsAPOL, pMensagem) Then
            'Se n�o encontrou, n�o h� o que fazer.
            GoTo SAIDA
        End If


        If Not rsAPOL.EOF Then
            Dim lobjTelefones As New clsA46V736A
            Dim lobjstcTelefones As New stcA46V736B

            Set lobjTelefones = New clsA46V736A
            Set lobjTelefones.M46V999 = M46V999

            Set lobjTelefones.ConexaoBD = lobjPedido.ConexaoBD
            Do While Not rsAPOL.EOF
                Set lobjstcTelefones = New stcA46V736B
                Set lobjstcTelefones.M46V999 = M46V999

                With lobjstcTelefones
                    .nosso_numero = lobjstcPedido.nosso_numero
                    .NUM_PI = lobjstcPedido.NUM_PI
                    .Tip_Telefone = M46V999.gfTratarNulo(rsAPOL("Tip_Telefone"), tpoNumerico)
                    .DDD_Telefone = M46V999.gfTratarNulo(rsAPOL("DDD_Telefone"), tpoNumerico)
                    .Numero_Telefone = M46V999.gfTratarNulo(rsAPOL("Numero_Telefone"), tpoNumerico)
                    .Ramal_Telefone = M46V999.gfTratarNulo(rsAPOL("Ramal_Telefone"), tpoNumerico)
                    .Obs_Telefone = M46V999.gfTratarNulo(rsAPOL("Obs_Telefone"), tpoString)
                    .Ind_Principal = M46V999.gfTratarNulo(rsAPOL("Ind_Principal"), tpoNumerico)
                End With
                Call lobjTelefones.mfIncluir(lobjstcTelefones, lPropTabela, "", tagDynaTrace)
                rsAPOL.MoveNext
            Loop
        End If


    Else
        pMensagem = "FGE0089 - Dados de comum n�o localizados."
    End If

    Set pStcPed = lobjstcPedido

    mfGraRenComum = True
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)

SAIDA:
    GoSub FechaObj
    Exit Function

ControleErro:
    Call M46V999.gpGraLog(1, "gfGraRenComum ERRO - " & Err & " " & Err.Description & " " & pMensagem)
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "gfGraRenComum ERRO - " & Err & " " & Err.Description & " " & pMensagem, 2, Class_Name, methodName)
    GoTo SAIDA

FechaObj:
    Set rsAPOL = Nothing
    Set lobjPedido = Nothing
    Set lobjstcPedido = Nothing
    Set lobjCorr = Nothing
    Set lObjstcCorr = Nothing
    Set lobjCoss = Nothing
    Set lobjstcCoss = Nothing
    Set lObjEnd = Nothing
    Set lObjstcEnd = Nothing
    Set lobjCtrlEmis = Nothing
    Set lobjstcCtrlEmis = Nothing
    Set lobjJuros = Nothing
    Set lobjAuto = Nothing
    Set lobjEspec = Nothing
    Set lobjstcEspec = Nothing
    Set clsA99V001 = Nothing
    Return
End Function


Private Function mfGraRenItem(prstItem As clsYasRecordSet, pNosNumero As String, _
                              pBdP00APOL As ADODB.Connection, pMensagem As String, _
                              pNumPed As Long, pCod_Dizer As Integer, pNumCotac_LMI As String, _
                              ByVal pNumApolice As String, pStcPed As stcA46V702B, tagDynaTrace As String) As Boolean

    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "mfGraRenItem"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)

'Fun��o: Grava os registros de comum com os dados da ap�lice (Renova��o Yasuda).
'Par�metros de entrada :prstItem.........Registro de item(ap�lice)
'                       pNosNumero.......Nosso n�mero.
'                       pBdP00APOL.......Arquivo de conex�o
'Par�metro de sa�da:    pMensagem........Mensagem de erro.

    Dim lSql As String        'Consulta SQL.
    Dim rsAPOL As clsYasRecordSet    'Registro.
    Dim rsNovaFranquia As clsYasRecordSet
'[FRH]->
'Conferir coberturas na tabela POOMULTI:TAB_COB_RELACAO
    Dim lSql_buscaCob As String
    Dim rs_buscaCob As clsYasRecordSet
    Dim bdP00MULTI As ADODB.Connection
    '<-[FRH]

    Dim lobjPedLoc As clsA46V708A         'Objeto da classe clsA46V708 (TAB_PED_LOC).
    Dim lobjstcPedLoc As stcA46V708B            'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC).
    Dim lObjClau As clsA46V703A       'Objeto da classe clsA46V703A (TAB_PED_CLAU).
    Dim lobjstcClau As stcA46V703B          'Objeto da classe de estrutura stcA46V703B (TAB_PED_CLAU).
    Dim lobjCobert As clsA46V704A         'Objeto da classe clsA46V704A (TAB_PED_COBERT).
    Dim lobjstcCobert As stcA46V704B            'Objeto da classe de estrutura stcA46V704B (TAB_PED_COBERT).
    Dim lobjInsp As clsA46V717A       'Objeto da classe clsA46V717A (TAB_PED_INSP).
    Dim lObjstcInsp As stcA46V717B          'Objeto da classe de estrutura stcA46V717B (TAB_PED_INSP).
    Dim lNumero As String
    Dim lComplemento As String
    Dim lPos As Long
    Dim lNomeTab As String
    Dim lValorRisco As Double
    Dim lObjPedFator As clsA46V720A
    Dim lobjstcPedFator As stcA46V720B
    Dim lPropTabela As String
    Dim lTemAssistencia As Boolean
    Dim lrsFranquia As clsYasRecordSet
    Dim lobjMultFraq As clsA46V738A    'Objeto da classe clsA46V738A (TAB_PED_COBERT_FRANQUIA)
    Dim lobjstcMultFraq As stcA46V738B    'Objeto da classe de estrutura stcA46V738B (TAB_PED_COBERT_FRANQUIA)
    Dim lobjPedLocBem As clsA46V715A
    Dim lobjstcPedLocBem As stcA46V715B
    Dim lobjPedLocBicicleta As clsA46V744A
    Dim lobjstcPedLocBicicleta As stcA46V744B
    Dim lObjPedCobertFranquia As clsA46V738A
    Dim lObjstcPedCobertFranquia As stcA46V738B
    Dim lObsFranquia As String
    Dim lobjPedAllRisks As clsA46V747A
    Dim lobjstcPedAllRisks As stcA46V747B
    Dim lobjPedObrasArte As clsA46V748A
    Dim lobjstcPedObrasArte As stcA46V748B

    On Error GoTo ControleErro
    Call M46V999.gfObtSYASMULTINI(lPropTabela)

    If M46V999.gPrograma = "D046O003" And pNumCotac_LMI = "" Then
        lNomeTab = "Syasportal.dbo.TAB_APOL"
    ElseIf M46V999.gBaseLocal = True Then
        lNomeTab = "TAB_APOL"
    Else
'        If pNumCotac_LMI <> "" Then    '1100177 - Renova��o Apolice LMI
'            lNomeTab = "RAMOSDIVERSOS.DBO.TAB_PED"  '1100177 - Renova��o Apolice LMI
'        Else
'            If pCod_Dizer = 0 Then
'                lNomeTab = "GED.DBO.TAB_PED"
'            Else
                lNomeTab = "TAB_APOL"
'            End If
'        End If
    End If

    If M46V999.gCodProduto = 0 Then
        M46V999.gCodProduto = pStcPed.Cod_Produto
    End If

    Set lobjPedLoc = New clsA46V708A
    Set lobjPedLoc.M46V999 = M46V999

    Set lobjstcPedLoc = New stcA46V708B
    Set lobjstcPedLoc.M46V999 = M46V999

    With lobjstcPedLoc
        .nosso_numero = pNosNumero
        .Num_Ped = pNumPed
        If M46V111.gNum_Pi = 0 Then
            .NUM_PI = Right(Year(Date), 2) & Format(pNumPed, "00000000")
        Else
            .NUM_PI = M46V111.gNum_Pi
        End If
        
        If M46V999.gfTratarNulo(prstItem("Cod_Org"), tpoNumerico) = 2 Then
            .Tip_Emissao = e_TipEmissao_RenCongenere
            .Ren_Con_Seguradora = 572
            .Ren_Con_Apolice = M46V999.gfTratarNulo(prstItem("Num_Apol"), tpoNumerico)
            .Ren_Con_Item = M46V999.gfTratarNulo(prstItem("Num_Item"), tpoNumerico)
            .Ind_Sinistro = M46V999.gfTratarNulo(prstItem("IND_SINISTRO"), tpoNumerico)
        End If
        
        .Cod_Ramo = M46V999.gfTratarNulo(prstItem("Cod_Ramo"), tpoNumerico)
        .Num_Item = M46V999.gfTratarNulo(prstItem("Num_item"), tpoNumerico)
        .Cod_Iden_Tarif = M46V999.gCodProduto
        .Cod_Tip_Logr_Risco = M46V999.gfTratarNulo(prstItem("Cod_Tip_Logr_Risco"), tpoString)
        .Nom_Logr_Risco = M46V999.gfTratarNulo(prstItem("Nom_Logr_Risco"), tpoString)
        lComplemento = M46V999.gfTratarNulo(prstItem("Dsc_Compl_Risco"), tpoString)
        lNumero = M46V999.gfTratarNulo(prstItem("NUm_LOgr_Risco"), tpoString)
        
        If Not M46V999.gfPreenchido(lNumero) Then
            If Left(lComplemento, 2) = "N." Then
                lNumero = Mid(lComplemento, 3, InStr(1, lComplemento, "-") - 3)
                lComplemento = Mid(lComplemento, InStr(1, lComplemento, "-") + 1)
            Else
                If IsNumeric(lComplemento) Then
                    lNumero = lComplemento
                    lComplemento = ""
                Else
                    lPos = InStr(1, lComplemento, " ")
                    If lPos > 0 Then
                        If IsNumeric(Left(lComplemento, lPos)) Then
                            lNumero = Left(lComplemento, lPos)
                            lComplemento = Mid(lComplemento, lPos + 1)
                        Else
                            lNumero = 0
                        End If
                    Else
                        lNumero = 0
                    End If
                End If
            End If
         End If
            
        .Dsc_Compl_Risco = Left(lComplemento, 20)
        .Nom_Bairro_Risco = M46V999.gfTratarNulo(prstItem("Nom_Bairro_Risco"), tpoString)
        .Nom_Cid_Risco = M46V999.gfTratarNulo(prstItem("Nom_Cid_Risco"), tpoString)
        .Sig_UF_Risco = M46V999.gfTratarNulo(prstItem("Sig_UF_Risco"), tpoString)
        .Num_Cep_Risco = M46V999.gfTratarNulo(prstItem("Num_CEP_Risco"), tpoNumerico)
        .Per_Desc_Apar = M46V999.gfTratarNulo(prstItem("Per_Desc_Apar"), TpoValor)

        If M46V111.gCod_plano >= 2021 And M46V111.gCod_plano <= 2031 Then
            .Per_Desc_Progres = 0
        ElseIf M46V111.gCod_plano >= 2061 And M46V111.gCod_plano <= 2100 Then
            .Per_Desc_Progres = 0
            'Ficha 901673
        ElseIf M46V111.gCod_plano >= 2111 And M46V111.gCod_plano <= 2120 Then
            .Per_Desc_Progres = 0
        Else
            .Per_Desc_Progres = M46V999.gfTratarNulo(prstItem("Per_Desc_Progres"), TpoValor)
        End If
        .Per_IS_Vr = M46V999.gfTratarNulo(prstItem("Per_IS_Vr"), TpoValor)
        .Val_Coef_IS_Vr = M46V999.gfTratarNulo(prstItem("Val_Coef_IS_Vr"), TpoValor)
        If pCod_Dizer = 0 Then
            If pStcPed.objstcCtrlEmis.cod_origem <> 2 Then
                .Tip_Emissao = e_TipEmissao_RenYasuda
            Else
                .Tip_Emissao = e_TipEmissao_RenCongenere
            End If
        Else
            .Tip_Emissao = e_TipEmissao_Endosso
        End If
        
        'nova tarifa 1100
'        If M46V999.gOrigemChamada = e_Origem_Web Then
'            M46V999.gCorrecaoMonetariaIPC = True
'        Else
'            M46V999.gCorrecaoMonetariaIPC = False
'        End If
        
'        If pStcPed.Cod_Estip > 2000 Then
'            M46V999.gCorrecaoMonetariaIPC = False
'        End If
        
'        If .Tip_Emissao = e_TipEmissao_RenYasuda And .Cod_Iden_Tarif >= 1100 And M46V999.gCorrecaoMonetariaIPC Then
'            .Val_Vr = Round(M46V999.gfTratarNulo(prstItem("Val_Vr"), TpoValor) * (1 + (pStcPed.objstcCtrlEmis.Indice_IPC / 100)), 2)
'        Else
            .Val_Vr = M46V999.gfTratarNulo(prstItem("Val_Vr"), TpoValor)
'        End If
        lValorRisco = .Val_Vr

        M46V999.gNum_Apolice = prstItem("num_apol")
        If pStcPed.Cod_Produto >= 1080 And _
           pStcPed.Cod_Dizer = 0 And _
           pStcPed.objstcCtrlEmis.cod_origem <> 2 And _
           M46V999.gTarifa_Apolice < M46V999.TarifaPrecificacaoMaritima Then
            If pStcPed.Cod_Estip < 2000 Then 'N�O FAZER DE-> PARA DE PLANINHO
                            
                M46V999.gCodProduto = pStcPed.Cod_Produto
                M46V999.gCod_Ramo = pStcPed.Cod_Ramo
                M46V999.M46V111.gCod_plano = pStcPed.Cod_Estip
            
                Call M46V999.mfCarregarDePara
            
                'Residencial Supremo
                '2017-10-30: Correcao de TipoEmissao
                'If .Val_Vr > 1000000# And pStcPed.Cod_Ramo = 113 And M46V999.LiberaPlanoResidencialSupremo(pStcPed.Cod_Estip, pStcPed.ColecaoPedLoc(1).Tip_Emissao) Then
                If .Val_Vr > 2000000# And pStcPed.Cod_Ramo = 113 And M46V999.LiberaPlanoResidencialSupremo(pStcPed.Cod_Estip, .Tip_Emissao) Then
                    .Tab_Clasf = 30 ' Residencial Supremo
                    M46V999.M46V111.gCod_plano = .Tab_Clasf
                Else
                    .Tab_Clasf = pStcPed.Cod_Estip
                End If

                .Cod_Clasf = M46V999.gfBuscarDePara(prstItem("Cod_Clasf"), 2, prstItem("Cod_Compl_Clasf"), False) '1000871 - Inclus�o de Org�o Categoria Conselho
                
                If Len(Trim(.Cod_Clasf)) < 3 And .Cod_Ramo = 114 Then ' Para renova��o sen�o encontrar o de->Para obriga o usuario selecionar
                    .Cod_Clasf = ""
                    .Cod_Compl_Clasf = ""
                Else
                    .Cod_Compl_Clasf = M46V999.gfBuscarDePara(prstItem("Cod_Clasf"), 2, prstItem("Cod_Compl_Clasf"), True)    '1000871 - Inclus�o de Org�o Categoria Conselho
                End If
            Else
                .Tab_Clasf = M46V999.gfTratarNulo(prstItem("Tab_Clasf"), tpoNumerico)
                .Cod_Clasf = M46V999.gfTratarNulo(prstItem("Cod_Clasf"), tpoString)  '1000871 - Inclus�o de Org�o Categoria Conselho
                .Cod_Compl_Clasf = M46V999.gfTratarNulo(prstItem("Cod_Compl_Clasf"), tpoString)    '1000871 - Inclus�o de Org�o Categoria Conselho
            End If
        Else
            .Tab_Clasf = M46V999.gfTratarNulo(prstItem("Tab_Clasf"), tpoNumerico)
            .Cod_Clasf = M46V999.gfTratarNulo(prstItem("Cod_Clasf"), tpoString)  '1000871 - Inclus�o de Org�o Categoria Conselho
            .Cod_Compl_Clasf = M46V999.gfTratarNulo(prstItem("Cod_Compl_Clasf"), tpoString)    '1000871 - Inclus�o de Org�o Categoria Conselho
        End If
        
        'Nova Tarifa 1100
        If .Cod_Iden_Tarif >= 1100 And .Tab_Clasf < 2000 And M46V999.gRst_Tbm_De_Para.State <> 0 Then
            .Cod_Clasf = M46V999.gfBuscarDePara(prstItem("Cod_Clasf"), 2, prstItem("Cod_Compl_Clasf"), False)
            .Cod_Compl_Clasf = M46V999.gfBuscarDePara(prstItem("Cod_Clasf"), 2, prstItem("Cod_Compl_Clasf"), True)
        End If
        .Cod_Plano = .Tab_Clasf
        .Dsc_Clasf = M46V999.gfTratarNulo(prstItem("Dsc_Clasf"), tpoString)
        .Dsc_Clasf = ""
        .Dsc_Clasf = Left(.gNom_Class, 130) 'Tarifa 11.12
        .Cod_Constr = M46V999.gfTratarNulo(prstItem("Cod_Constr"), tpoNumerico)
        .Cod_User_Liber = M46V999.gNomUsuario6
        .Dat_Liber = Left(M46V999.gfForData(3), 8)
        .Hor_Liber = Mid(M46V999.gfForData(3), 9, 4)
        .Cod_Sit_item = e_SitItem_NaoCalc
        .Per_Bonus = M46V999.gfTratarNulo(prstItem("Per_Bonus"), tpoNumerico)
        .Num_Apol_Anterior = prstItem("Num_APol")
        
        'Ajuste na numera��o do item quando renova��o de item especifico
        If prstItem.RecordCount = 1 And .Tip_Emissao <> e_TipEmissao_Endosso Then
            If .Num_Item <> 1 Then
                .Num_Item_Anterior = .Num_Item
                .Num_Item = 1
            End If
        End If
        
        '11.50 2 - Cria��o de nova atividade Bicicleta Loja, Dep�sito e oficina.
        If (.Cod_Clasf = "145" Or .Cod_Clasf = "146") And .Cod_Ramo = 112 And pCod_Dizer = 0 And .Cod_Iden_Tarif >= 1150 Then
            .Cod_Clasf = "584"
            .Cod_Compl_Clasf = "01"
            .Dsc_Clasf = Empty
            .Dsc_Clasf = Left(.gNom_Class, 130)
        End If
        
        .Num_Item_Anterior = prstItem("Num_Item")
        If pNumCotac_LMI <> "" Then    '1100177 - Renova��o Apolice LMI
            .Num_Apol_Anterior = pNumApolice
        End If

        If lobjstcPedLoc.mBonus_Anterior = 0 And lobjstcPedLoc.BonusAlterado = False Then
            .Per_Bonus = lobjPedLoc.mfRetornaBonusApolice(lobjstcPedLoc, tagDynaTrace)
        Else
            .Per_Bonus = lobjstcPedLoc.mBonus_Anterior
        End If

        If pStcPed.objstcCtrlEmis.cod_origem <> 2 Then
        .Ind_Sinistro = lobjstcPedLoc.Apol_Ant_IND_SINISTRO
        End If
        If .Ind_Sinistro = 0 Then .Ind_Sinistro = e_SimNao_Nao
        If pCod_Dizer = 0 Then

            'Aplica o bonus olhando o sinistro
            If .Ind_Sinistro = e_SimNao_Sim Then
                .Per_Bonus = .Per_Bonus - 5
            Else
                .Per_Bonus = .Per_Bonus + 5
            End If
            If .Per_Bonus < 0 Then .Per_Bonus = 0

            If lobjstcPedLoc.Cod_Iden_Tarif < 1080 Then
                'Soma mais 5 da fidelidade.
                .Per_Bonus = .Per_Bonus + 5
            End If

            'If .Per_Bonus > 30 Then .Per_Bonus = 30
            If .Per_Bonus > 25 Then .Per_Bonus = 25

            '            'Aplico a regra da fidelidade
            '            If .Cod_Ramo = e_CodRamo_Empresarial Then
            '
            '                If M46V999.gCodProduto >= 942 Then
            '                    'Soma mais 5 da fidelidade.
            '                    .Per_Bonus = .Per_Bonus + 5
            '                End If
            '                If .Per_Bonus < 0 Then .Per_Bonus = 0
            '                If .Per_Bonus > 30 Then .Per_Bonus = 30
            '
            '            ElseIf .Cod_Ramo = e_CodRamo_Residencial Then
            '                If M46V999.gCodProduto >= 946 Then
            '                    'Soma mais 5 da fidelidade.
            '                    .Per_Bonus = .Per_Bonus + 5
            '                    If .Per_Bonus > 30 Then .Per_Bonus = 30
            '                Else
            '                    If .Per_Bonus > 25 Then .Per_Bonus = 25
            '                End If
            '                If M46V111.gCod_Plano >= 2001 And M46V111.gCod_Plano <= 2031 Then
            '                    .Per_Bonus = 0
            '                End If
            '            'Ficha 901673
            '                If (M46V111.gCod_Plano >= 2051 And M46V111.gCod_Plano <= 2100) Or (M46V111.gCod_Plano >= 2111 And M46V111.gCod_Plano <= 2120) Then
            '                    .Per_Bonus = 0
            '                End If
            '            End If
        End If

        ' Ficha CR1425 - Planos N�O podem ter desconto de experi�ncia
        If M46V111.gCod_plano >= 2001 _
        And M46V111.gCod_plano <> 2300 _
        And pStcPed.getPlano.EhPlanoSegmentado = False Then
            .Per_Bonus = 0
        End If
        ' Ficha CR1425 - Fim


        'Ficha 900550
        M46V111.gDesc_Bonus = .Per_Bonus

        .Qtd_Empreg = M46V999.gfTratarNulo(prstItem("Qtd_Empreg"), tpoNumerico)
        .Cod_Iden_Segur_01_V = M46V999.gfTratarNulo(prstItem("Cod_Iden_Segur_01"), tpoString)
        .Cod_Iden_Segur_02_A = M46V999.gfTratarNulo(prstItem("Cod_Iden_Segur_02"), tpoString)
        .Cod_Iden_Segur_03_P = M46V999.gfTratarNulo(prstItem("Cod_Iden_Segur_03"), tpoString)
        .Cod_Iden_Segur_04_I = M46V999.gfTratarNulo(prstItem("Cod_Iden_Segur_04"), tpoString)                    'Inexist�ncia de terreno baldio,  im�vel desocupado   ou em constru��o na vizinhan�a do risco.: S ou N
        .Cod_Iden_Segur_05_G = M46V999.gfTratarNulo(prstItem("Cod_Iden_Segur_05"), tpoString)                    'Grades met�licas de prote��o em todas as janelas e portas de vidros que d�o acesso ao interior do im�vel desde que tamb�m possuam alarme.: S ou N
        .Per_Tot_Segur = .gPer_Tot_Segur
        .Cod_Iden_Protec_01_E = M46V999.gfTratarNulo(prstItem("Cod_Iden_Protec_01"), tpoString)
        .Cod_Iden_Protec_02_H = M46V999.gfTratarNulo(prstItem("Cod_Iden_Protec_02"), tpoString)
        .Cod_Iden_Protec_03_S = M46V999.gfTratarNulo(prstItem("Cod_Iden_Protec_03"), tpoString)
        .Cod_Iden_Protec_04_D = M46V999.gfTratarNulo(prstItem("Cod_Iden_Protec_04"), tpoString)                   'Indicador de Detec��o e alarmes: S ou N.
        .Per_Tot_Protec = .gPer_Tot_Protec
        If M46V111.gCod_plano > 1010 Then
            .Cod_Produto = M46V999.gfTratarNulo(prstItem("Cod_Produto"), tpoNumerico)
        Else
            If .Cod_Iden_Tarif >= 1080 Then
                .Cod_Produto = pStcPed.Cod_Estip
            Else
                .Cod_Produto = 1
            End If
        End If

        .Cod_Regiao = M46V999.gfTratarNulo(prstItem("Cod_Regiao"), tpoNumerico)
        
        If pStcPed.objstcCtrlEmis.cod_origem <> 2 Then
        .Ind_Sinistro = lobjstcPedLoc.Apol_Ant_IND_SINISTRO
        End If
        If .Ind_Sinistro = 0 Then .Ind_Sinistro = e_SimNao_Nao
        .Ind_Sin_Inspec = M46V999.gfTratarNulo(prstItem("Ind_Sin_Inspec"), tpoNumerico)
        .Cod_DizerItem = pCod_Dizer
        ''Ficha 1000305
        .Tip_Cobert = M46V999.gfTratarNulo(prstItem("Tip_Cobert"), tpoNumerico)
        '        M46V999.gTipCobertura = .Tip_Cobert
        .Possui_Elevador = M46V999.gfTratarNulo(prstItem("POSSUI_ELEVADOR"), tpoNumerico)
        .BonusAlterado = M46V999.gfTratarNulo(prstItem("BonusAlterado"), TpoBoolean)

        .Ind_VerbaSep = M46V999.gfTratarNulo(prstItem("Ind_VerbaSep"), tpoNumerico)
        
        'nova tarifa 1100
'        If .Tip_Emissao = e_TipEmissao_RenYasuda And .Cod_Iden_Tarif >= 1100 And M46V999.gCorrecaoMonetariaIPC Then
'            .Val_Vr_Lc = Round(M46V999.gfTratarNulo(prstItem("Val_Vr_Lc"), TpoValor) * (1 + (pStcPed.objstcCtrlEmis.Indice_IPC / 100)), 2)
'        Else
            .Val_Vr_Lc = M46V999.gfTratarNulo(prstItem("Val_Vr_Lc"), tpoNumerico)
            .Val_LMG_Lc = M46V999.gfTratarNulo(prstItem("Val_LMG_Lc"), TpoValor)
'        End If
        
        .Per_IS_Vr_Lc = M46V999.gfTratarNulo(prstItem("Per_IS_Vr_Lc"), tpoNumerico)
        .Coef_IS_Vr_Lc = M46V999.gfTratarNulo(prstItem("Coef_IS_Vr_Lc"), tpoNumerico)

        .Tip_Plano_Vida = M46V999.gfTratarNulo(prstItem("Tip_Plano_Vida"), tpoNumerico)
        .Cod_Plano_Cobertura = M46V999.gfTratarNulo(prstItem("Cod_Plano_Cobertura"), tpoNumerico)
        .Qtd_Vidas = M46V999.gfTratarNulo(prstItem("Qtd_Vidas"), tpoNumerico)
        .Val_LMG_Vida = M46V999.gfTratarNulo(prstItem("Val_LMG_Vida"), tpoNumerico)
        .Ind_Aux_Funeral = M46V999.gfTratarNulo(prstItem("Ind_Aux_Funeral"), tpoNumerico)
        .Val_LMG_Funeral = M46V999.gfTratarNulo(prstItem("Val_LMG_Funeral"), tpoNumerico)
        .Ind_Inc_Conjuge = M46V999.gfTratarNulo(prstItem("Ind_Inc_Conjuge"), tpoNumerico)
        .Val_Aux_Cesta = M46V999.gfTratarNulo(prstItem("Val_Aux_Cesta"), tpoNumerico)

        Set .ObjPai = pStcPed
    End With

    If Not lobjPedLoc.mfIncluir(lobjstcPedLoc, lPropTabela, "", tagDynaTrace) Then
        pMensagem = lobjPedLoc.UltimoErro
        GoTo SAIDA
    End If
    
    ' Benefici�rios
    

    'Clausulas
    lTemAssistencia = False
    lSql = " Select * from " & lNomeTab & "_CLAU WITH (NOLOCK) "
    If M46V999.gPrograma <> "D046O003" Or pNumCotac_LMI <> "" Then
        If M46V999.gBaseLocal = False And pBdP00APOL Like "*Provider=SQLOLEDB.*" Then
            lSql = lSql & " inner join " & IIf(pNumCotac_LMI = "", "GED", "RAMOSDIVERSOS") & ".DBO.TAB_PED  as ctrl WITH (NOLOCK) on " & lNomeTab & "_CLAU.num_apol = ctrl.num_apol"
            lSql = lSql & " and ctrl.num_endo = " & lNomeTab & "_CLAU.num_endo "
            lSql = lSql & " and ctrl.num_PI = " & lNomeTab & "_CLAU.num_PI "
            If pNumCotac_LMI <> "" Then    '1100177 - Renova��o Apolice LMI
                lSql = lSql & " and ctrl.NOSSO_NUMERO = " & lNomeTab & "_CLAU.NOSSO_NUMERO "
            End If
        End If
    End If

    If pNumCotac_LMI <> "" Then    '1100177 - Renova��o Apolice LMI
        lSql = lSql & " Where " & lNomeTab & "_CLAU.NOSSO_NUMERO = " & M46V999.gfFormatarCamposSQL(pNumCotac_LMI)
    Else
        lSql = lSql & " Where " & lNomeTab & "_CLAU.Num_Apol = " & prstItem("Num_Apol")
        lSql = lSql & " and " & lNomeTab & "_CLAU.num_Endo = " & prstItem("NUm_Endo2")
    End If
    lSql = lSql & " and " & lNomeTab & "_CLAU.num_item = " & prstItem("Num_Item")
    If Not M46V999.gfObtRegistro(pBdP00APOL, lSql, rsAPOL, pMensagem) Then
        GoTo ControleErro
    End If

    Set lObjClau = New clsA46V703A
    Set lObjClau.M46V999 = M46V999

    Set lObjClau.ConexaoBD = lobjPedLoc.ConexaoBD
    Do While Not rsAPOL.EOF
        Set lobjstcClau = New stcA46V703B
        Set lobjstcClau.M46V999 = M46V999

        With lobjstcClau
            .Cod_Ramo = M46V999.gfTratarNulo(rsAPOL("Cod_Ramo"), tpoNumerico)
            .nosso_numero = pNosNumero
            .Num_Ped = pNumPed
            If M46V111.gNum_Pi = 0 Then
                .NUM_PI = Right(Year(Date), 2) & Format(pNumPed, "00000000")
            Else
                .NUM_PI = M46V111.gNum_Pi
            End If
            .Cod_Clau = M46V999.gfTratarNulo(rsAPOL("Cod_Clau"), tpoNumerico)
            .Num_Item = lobjstcPedLoc.Num_Item
        End With
        If Not lObjClau.mfIncluir(lobjstcClau, lPropTabela, "", tagDynaTrace) Then
            pMensagem = lObjClau.UltimoErro
            GoTo SAIDA
        End If

        If lobjstcClau.Cod_Clau = 400 Or lobjstcClau.Cod_Clau = 401 Or lobjstcClau.Cod_Clau = 402 Then
            lTemAssistencia = True
        End If
        rsAPOL.MoveNext
    Loop

    'Adicionado a conex�o para acesso a base de dados nas consultas e consistencias.
    Set lobjstcPedLoc.M46V999 = M46V999
    
    If lTemAssistencia = False _
        And lobjstcPedLoc.gPodeOptarAss24h And _
        lobjstcClau.Cod_Ramo <> e_CodRamo_Condominio Then '8000048288, Informa��o da clausula 400
    
        Set lobjstcClau = New stcA46V703B
        Set lobjstcClau.M46V999 = M46V999

        With lobjstcClau
            .Cod_Ramo = lobjstcPedLoc.Cod_Ramo
            .nosso_numero = pNosNumero
            .Num_Ped = pNumPed
            If M46V111.gNum_Pi = 0 Then
                .NUM_PI = Right(Year(Date), 2) & Format(pNumPed, "00000000")
            Else
                .NUM_PI = M46V111.gNum_Pi
            End If
            .Cod_Clau = 400
            .Num_Item = lobjstcPedLoc.Num_Item
        End With
        If Not lObjClau.mfIncluir(lobjstcClau, lPropTabela, "", tagDynaTrace) Then
            pMensagem = lObjClau.UltimoErro
            GoTo SAIDA
        End If
    End If
    'COberturas
    lSql = " Select cob.*, cobC.TextoFranquia, cobC.Val_Fator_Franq, cobC.Compl_Desc_Cobertura "
    If M46V999.gBaseLocal = False And pBdP00APOL Like "*Provider=SQLOLEDB.*" Then
        lSql = lSql & " , Num_Meses_Pi, Val_Risco_Cob, Val_Pr_Orig  "
    End If
    lSql = lSql & " from " & lNomeTab & "_COBERT  as cob WITH (NOLOCK) "

    If (M46V999.gBaseLocal = False And pBdP00APOL Like "*Provider=SQLOLEDB.*" And M46V999.gPrograma <> "D046O003") Or pNumCotac_LMI <> "" Then
        If M46V999.gfTratarNulo(prstItem("Cod_Org"), tpoNumerico) < 2 Then
                lSql = lSql & " inner join " & IIf(pNumCotac_LMI = "", "GED", "RAMOSDIVERSOS") & ".DBO.tab_ped as ped on "
        Else
            ' Para renova��o maritima
            lSql = lSql & " inner join tab_apol as ped on "
        End If
        'lSql = lSql & " inner join " & lNomeTab & " as ped on "
        lSql = lSql & " ped.num_apol = cob.num_apol and "
        'lSql = lSql & " ped.num_endo = cob.num_endo and "
        lSql = lSql & " ped.num_PI = cob.num_PI"
        If pNumCotac_LMI <> "" Then    '1100177 - Renova��o Apolice LMI
            lSql = lSql & " and ped.NOSSO_NUMERO = cob.NOSSO_NUMERO"
        End If
    End If

    If M46V999.gBaseLocal = False And pBdP00APOL Like "*Provider=SQLOLEDB.*" Then
        lSql = lSql & " left join " & lNomeTab & "_COBERT_COMPL  as cobC on "
        lSql = lSql & " cobC.num_apol = cob.num_apol and "
        lSql = lSql & " cobC.num_endo = cob.num_endo and "
        lSql = lSql & " cobC.num_PI = cob.num_PI and "
        lSql = lSql & " cobC.num_Item = cob.num_item and "
        lSql = lSql & " cobC.cod_cobert = cob.cod_cobert"
    End If
    If pNumCotac_LMI <> "" Then    '1100177 - Renova��o Apolice LMI
        lSql = lSql & " and cobC.NOSSO_NUMERO = cob.NOSSO_NUMERO  "
    End If
    If pNumCotac_LMI <> "" Then    '1100177 - Renova��o Apolice LMI
        lSql = lSql & " Where ped.NOSSO_NUMERO = " & M46V999.gfFormatarCamposSQL(pNumCotac_LMI)
    Else
        lSql = lSql & " Where cob.Num_Apol = " & prstItem("Num_Apol")
        'lSql = lSql & " and   cob.num_Endo = " & prstItem("NUm_Endo2")
    End If
    lSql = lSql & " and  cob.num_item = " & prstItem("Num_Item")
    lSql = lSql & " and  cob.val_is > 0"
    If M46V999.gCodProduto >= 1122 And pCod_Dizer = 0 Then  'Tarifa 11.22
        lSql = lSql & " and  NOT (cob.Cod_Ramo = 114 and cob.Tab_Cobert = 11 and cob.Cod_Cobert = 9636)"
    End If
    
    If prstItem("Cod_Ramo") = e_CodRamo_Riscos_diversos Then
        lSql = lSql & " order by cob.tab_cobert, cob.cod_cobert "
    End If

    Set lobjCobert = New clsA46V704A
    Set lobjCobert.M46V999 = M46V999

    Set lobjCobert.ConexaoBD = lobjPedLoc.ConexaoBD
    '[FRH]->
    If Not M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, bdP00MULTI, pMensagem) Then
        GoTo ControleErro
    End If

    If Not M46V999.gfObtRegistro(pBdP00APOL, lSql, rsAPOL, pMensagem) Then
        GoTo ControleErro
    End If

    '<-[FRH]
    If rsAPOL.RecordCount = 0 Then
        pMensagem = "Processo abortado, favor verificar os dados da ap�lice."
        GoTo ControleErro
    End If
    
    Do While Not rsAPOL.EOF
        Set lobjstcCobert = New stcA46V704B
        Set lobjstcCobert.M46V999 = M46V999

        With lobjstcCobert
            .Cod_Ramo = M46V999.gfTratarNulo(rsAPOL("Cod_Ramo"), tpoNumerico)
            .nosso_numero = pNosNumero
            .Num_Ped = pNumPed
            If M46V111.gNum_Pi = 0 Then
                .NUM_PI = Right(Year(Date), 2) & Format(pNumPed, "00000000")
            Else
                .NUM_PI = M46V111.gNum_Pi
            End If
            .Num_Item = lobjstcPedLoc.Num_Item

            '[FRH]->
            'Rotina para conferir altera��es de c�digos de cobertura,
            'na *POOMULTI.TAB_COB_RELACAO* existe um de-para de cobs.anteriores e atuais.

            '[Valor original Cod_Cobert da TAB_APOL]
            If pStcPed.Cod_Produto >= 1080 And pStcPed.Cod_Estip < 2000 Then
                .Cod_Cobert = M46V999.gfBuscarDePara(M46V999.gfTratarNulo(rsAPOL("cod_cobert"), tpoNumerico), 1)
            Else
                .Cod_Cobert = M46V999.gfTratarNulo(rsAPOL("cod_cobert"), tpoNumerico)
            End If

            lSql_buscaCob = "SELECT COD_COB_ATUAL FROM TAB_COB_RELACAO WITH (NOLOCK) WHERE COD_COB_ANTERIOR=" & M46V999.gfTratarNulo(rsAPOL("Cod_Cobert"), tpoNumerico)
            If M46V999.gfObtRegistro(bdP00MULTI, lSql_buscaCob, rs_buscaCob, pMensagem) Then
                If Not (rs_buscaCob.BOF And rs_buscaCob.EOF) Then
                    .Cod_Cobert = M46V999.gfTratarNulo(rs_buscaCob("cod_cob_atual"), tpoNumerico)
                End If
            End If
            '<-[FRH]

            If .Cod_Ramo = e_CodRamo_Riscos_diversos Then
                .Tab_Cobert = M46V999.gfTratarNulo(rsAPOL("Tab_Cobert"), tpoNumerico)
                M46V111.gCod_plano = .Tab_Cobert
            End If
            If M46V111.gCod_plano > 1010 Then
                .Tab_Cobert = M46V111.gCod_plano
            Else
                .Tab_Cobert = pStcPed.Cod_Estip
            End If
                        
            'CR93893 Para renova��o deve pegar o PI_Meses padr�o
            If pCod_Dizer > 0 Then
                .Num_Meses_PI = M46V999.gfTratarNulo(rsAPOL("Num_Meses_PI"), tpoNumerico)
            Else
                .Num_Meses_PI = M46V999.gfBuscaPerIndenitarioDefault(.Cod_Cobert, .Tab_Cobert)
            End If

            'Fichas 901282 e 901346
            Select Case M46V111.gCod_plano
                '1000744 - Item 4 - Plano Mercedes
            Case 2021 To 2031
                If M46V999.gCodProduto < 992 Then
                    If lobjstcPedLoc.Cod_Clasf = M46V999.gfBuscarDePara("001", 2) And lobjstcPedLoc.Cod_Compl_Clasf = M46V999.gfBuscarDePara("001", 2, "11") Then .Per_Desc_Coml = 33
                    If lobjstcPedLoc.Cod_Clasf = M46V999.gfBuscarDePara("001", 2) And lobjstcPedLoc.Cod_Compl_Clasf = M46V999.gfBuscarDePara("001", 2, "21") Then .Per_Desc_Coml = 20
                End If
            Case 2051 To 2060
                '1200903 - Tarifa - 1046 e 1047 - Item 13
                If M46V999.gCodProduto < 1046 Then
                    If lobjstcPedLoc.Cod_Clasf = M46V999.gfBuscarDePara("001", 2) And lobjstcPedLoc.Cod_Compl_Clasf = M46V999.gfBuscarDePara("001", 2, "11") Then .Per_Desc_Coml = 20
                    If lobjstcPedLoc.Cod_Clasf = M46V999.gfBuscarDePara("001", 2) And lobjstcPedLoc.Cod_Compl_Clasf = M46V999.gfBuscarDePara("001", 2, "21") Then .Per_Desc_Coml = 5
                End If
            Case 2061 To 2100
                If lobjstcPedLoc.Cod_Clasf = M46V999.gfBuscarDePara("001", 2) And lobjstcPedLoc.Cod_Compl_Clasf = M46V999.gfBuscarDePara("001", 2, "11") Then .Per_Desc_Coml = 20
                If lobjstcPedLoc.Cod_Clasf = M46V999.gfBuscarDePara("001", 2) And lobjstcPedLoc.Cod_Compl_Clasf = M46V999.gfBuscarDePara("001", 2, "21") Then .Per_Desc_Coml = 5
                'Ficha 901673
            Case 2111 To 2120
                If lobjstcPedLoc.Cod_Clasf = M46V999.gfBuscarDePara("001", 2) And lobjstcPedLoc.Cod_Compl_Clasf = M46V999.gfBuscarDePara("001", 2, "11") Then .Per_Desc_Coml = 20
                If lobjstcPedLoc.Cod_Clasf = M46V999.gfBuscarDePara("001", 2) And lobjstcPedLoc.Cod_Compl_Clasf = M46V999.gfBuscarDePara("001", 2, "21") Then .Per_Desc_Coml = 5
            Case Else
                If M46V111.gCod_plano > 2000 Or lobjstcPedLoc.Tip_Emissao = e_TipEmissao_Endosso Then
                        .Per_Desc_Coml = M46V999.gfTratarNulo(rsAPOL("Per_Desc_Coml"), tpoNumerico)
                End If
                
                '1200539 - Tarifa 1038 e 1039 Junho 2012 - Item 1
                Dim lCoef_Desc_Coml As Double
                If lobjstcPedLoc.Cod_Iden_Tarif >= 1038 Then
                    If M46V103.mfAjustesPontuais(e_Tip_Valor_DescComl, TpoValor, lCoef_Desc_Coml, pMensagem, lobjstcPedLoc, 0, tagDynaTrace) = False Then
                        GoTo ControleErro
                    End If
                    If .Per_Desc_Coml < lCoef_Desc_Coml And M46V111.gCod_plano > 2000 Then
                        .Per_Desc_Coml = Format(lCoef_Desc_Coml, "standard")
                    End If
                End If
            End Select

            'Ficha 900550
            '.Per_Desc_Exp = M46V999.gfTratarNulo(rsAPOL("Per_Desc_Exp"), tpoNumerico)

            .Per_Desc_Exp = M46V111.gDesc_Bonus


            'nova tarifa 1100
'            If lobjstcPedLoc.Tip_Emissao = e_TipEmissao_RenYasuda And lobjstcPedLoc.Cod_Iden_Tarif >= 1100 And M46V999.gCorrecaoMonetariaIPC Then
'                .Val_IS = Round(M46V999.gfTratarNulo(rsAPOL("Val_IS"), tpoNumerico) * (1 + (pStcPed.objstcCtrlEmis.Indice_IPC / 100)), 2)
'            Else
                .Val_IS = M46V999.gfTratarNulo(rsAPOL("Val_IS"), tpoNumerico)
'            End If
            
            .Val_Risco_Cob = .Val_IS

            If pCod_Dizer <> 0 Then
                If M46V111.gCod_plano > 1010 And .Cod_Ramo = e_CodRamo_Residencial Then
                    If InStr("3,6", pCod_Dizer) <> 0 Then
                        .Alterar_Cobert = "FIXO"
                    ElseIf InStr("2,4,10,11,18", pCod_Dizer) <> 0 Then
                        .Alterar_Cobert = "FIXA IS"
                    End If
                ElseIf .Cod_Ramo = e_CodRamo_Residencial Then
                    If InStr("3,6", pCod_Dizer) <> 0 Then
                        .Alterar_Cobert = "FIXO"
                    End If
                Else
                    If InStr("2,3,4,5,6,7,10,11,18", pCod_Dizer) <> 0 Then
                        .Alterar_Cobert = "FIXO"
                    End If
                End If
            End If

            '11/01/2010 - Ficha: 1000003 - franquia e valor de franquia - inicio
            'Call mpBuscaLimFranquia(lObjstcPedLoc, lobjstcCobert)
            If M46V999.gCodProduto < 1080 Then
                Dim lValFraqMinDefault As Double
                Dim lPerFranqDefault As Double
                Dim lFatorFranqDefault As Double
                Dim lValFraqMaxDefault As Double
                Dim TextoFranquia As String
                Dim Dummy1 As String
    
                Call M46V999.mpNovoBuscaLimFranquia(lobjstcPedLoc.Cod_Ramo, lobjstcPedLoc.Tab_Clasf, lobjstcPedLoc.ObjPai.Cod_Produto, _
                                                    lobjstcCobert.Cod_Cobert, lobjstcPedLoc.Cod_Clasf, lobjstcPedLoc.Cod_Compl_Clasf, _
                                                    lobjstcPedLoc.Cod_Constr, lobjstcPedLoc.Num_Cep_Risco, lobjstcPedLoc.ObjPai.Cod_Sequencia, lobjstcPedLoc.ObjPai.objstcCtrlEmis.Cod_Corr, _
                                                    lValFraqMinDefault, lPerFranqDefault, lFatorFranqDefault, lValFraqMaxDefault, TextoFranquia, Dummy1, lobjstcCobert.Num_Meses_PI)
    
    
                '11/01/2010 - Ficha: 1000003 - franquia e valor de franquia - fim
                'If lobjstcCobert.Val_Franq_Min < rsAPOL("val_franq_min") Or pCod_Dizer <> 0 Then
                '    lobjstcCobert.Val_Franq_Min = rsAPOL("val_franq_min")
                '    lobjstcCobert.Per_Franq = rsAPOL("per_franq")
                'End If
    
                If pCod_Dizer = 0 Then
                    lobjstcCobert.Val_Franq_Min = lValFraqMinDefault
                    lobjstcCobert.Per_Franq = lPerFranqDefault
                    lobjstcCobert.Val_Franq_Max = lValFraqMaxDefault
                    lobjstcCobert.Val_Fator_Franq = lFatorFranqDefault
                    lobjstcCobert.TextoFranquia = TextoFranquia
                Else
                    lobjstcCobert.Val_Franq_Min = rsAPOL("val_franq_min")
                    lobjstcCobert.Per_Franq = rsAPOL("per_franq")
                    lobjstcCobert.Val_Franq_Max = 0
                    lobjstcCobert.Val_Fator_Franq = 0
                    lobjstcCobert.TextoFranquia = M46V999.gfTextoFranquia(rsAPOL("val_franq_min"), rsAPOL("per_franq"), "", lobjstcCobert.Cod_Cobert, lobjstcCobert.Num_Meses_PI, _
                                                                          lobjstcCobert.Val_IS, lobjstcPedLoc)
                End If
            Else
                If pCod_Dizer = 0 Then

                    Call M46V999.mfFranqNovoLim(lobjstcCobert.Cod_Cobert, lobjstcCobert.Val_IS, _
                                                lobjstcPedLoc.Cod_Clasf, _
                                                lobjstcPedLoc.Cod_Compl_Clasf, _
                                                lobjstcPedLoc.Cod_Constr, _
                                                lobjstcPedLoc.Num_Cep_Risco, _
                                                lobjstcPedLoc.Tip_Cobert, _
                                                lobjstcCobert.Num_Meses_PI, lrsFranquia)
                    If lrsFranquia.RecordCount > 0 Then
                        lobjstcCobert.Val_Franq_Min = M46V999.gfTratarNulo(lrsFranquia("Val_Franq_Min"), tpoNumerico)
                        lobjstcCobert.Per_Franq = M46V999.gfTratarNulo(lrsFranquia("Perc_Franquia"), tpoNumerico)
                        lobjstcCobert.Val_Franq_Max = 0
                        lobjstcCobert.Val_Fator_Franq = M46V999.gfTratarNulo(lrsFranquia("Fator_Ajuste"), tpoNumerico)
                        lobjstcCobert.TextoFranquia = M46V999.mfMontaTexofranquia(M46V999.gfTratarNulo(lrsFranquia("Perc_Franquia"), tpoNumerico), M46V999.gfTratarNulo(lrsFranquia("Val_Franq_Min"), tpoNumerico), M46V999.gfTratarNulo(lrsFranquia("Dias_Franq_Min"), tpoNumerico), M46V999.gfTratarNulo(lrsFranquia("Texto"), tpoString), M46V999.gfTratarNulo(lrsFranquia("Compl"), tpoString), M46V999.gfTratarNulo(lrsFranquia!Tem_VideObs, tpoString))
                        lObsFranquia = M46V999.gfTratarNulo(lrsFranquia("Observacao_Franq"), tpoString)
                    End If
                    
                                        'CRQ000000440053
                    'Carregar Tab_Ped_Cobert_Franquia
                    Set lObjPedCobertFranquia = New clsA46V738A
                    Set lObjPedCobertFranquia.M46V999 = M46V999

                    Set lObjstcPedCobertFranquia = New stcA46V738B
                    Set lObjstcPedCobertFranquia.M46V999 = M46V999
                    
                    With lObjstcPedCobertFranquia
                        .Cod_Ramo = lobjstcCobert.Cod_Ramo
                        .Cod_Modalidade = lobjstcCobert.Tab_Cobert
                        .Cod_Tarif = M46V999.gCodProduto
                        .Cod_Cobert = lobjstcCobert.Cod_Cobert
                        .Cod_Evento = 0
                        .Num_Ped = lobjstcCobert.Num_Ped
                        .Num_Apol = lobjstcCobert.Num_Apol
                        .Num_Item = lobjstcCobert.Num_Item
                        .Num_Endosso = lobjstcCobert.Num_Endo
                        .NUM_PI = lobjstcCobert.NUM_PI
                        .nosso_numero = lobjstcCobert.nosso_numero
                        .Val_Min_Franq = lobjstcCobert.Val_Franq_Min
                        .Per_Franq = lobjstcCobert.Per_Franq
                        .Multiplo_Franquia = 1
                        .Texto_Observacoes = lObsFranquia
                        .Texto_Franquia = lobjstcCobert.TextoFranquia
                    End With
                    
                    If Not lObjPedCobertFranquia.mfIncluir(lObjstcPedCobertFranquia, "", "", "") Then
                        pMensagem = lObjPedCobertFranquia.UltimoErro
                        GoTo SAIDA
                    End If
                    
                    Set lObjstcPedCobertFranquia.ObjPai = lobjstcCobert
                    lobjstcCobert.ColecaoNovaFranquia.Add lObjstcPedCobertFranquia, lObjstcPedCobertFranquia.Class_ID
                       
                Else
                    lobjstcCobert.Val_Franq_Min = rsAPOL("val_franq_min")
                    lobjstcCobert.Per_Franq = rsAPOL("per_franq")
                    lobjstcCobert.Val_Franq_Max = 0
                    lobjstcCobert.Val_Fator_Franq = M46V999.gfTratarNulo(rsAPOL("Val_Fator_Franq"), tpoNumerico)
                    lobjstcCobert.TextoFranquia = "" & rsAPOL("TextoFranquia")
                    
'                    lobjstcCobert.TextoFranquia = M46V999.gfTextoFranquia(rsAPOL("val_franq_min"), _
'                                                                        rsAPOL("per_franq"), _
'                                                                        "", _
'                                                                        lobjstcCobert.Cod_Cobert, _
'                                                                        lobjstcCobert.Num_Meses_PI, _
'                                                                        lobjstcCobert.Val_IS, _
'                                                                        lobjstcPedLoc)
                End If
            End If

            If M46V999.gBaseLocal = False And pBdP00APOL Like "*Provider=SQLOLEDB.*" Then
                'CR93893 Pega o valor da apolice apenas para endosso
                If pCod_Dizer > 0 Then
                    lobjstcCobert.Num_Meses_PI = M46V999.gfTratarNulo(rsAPOL("Num_Meses_Pi"), tpoNumerico)
                End If
                
                If M46V999.gfTratarNulo(rsAPOL("Val_Risco_Cob"), tpoNumerico) <> 0 Then
                    lobjstcCobert.Val_Risco_Cob = M46V999.gfTratarNulo(rsAPOL("Val_Risco_Cob"), tpoNumerico)
                End If
            End If
            If Not IsNull(rsAPOL("Compl_Desc_Cobertura")) Then
                .Compl_Desc_Cobertura = rsAPOL("Compl_Desc_Cobertura")
            End If
        End With

        If Not lobjCobert.mfIncluir(lobjstcCobert, lPropTabela, "", False, False, tagDynaTrace) Then
            If UCase(lobjCobert.UltimoErro) Like "*DUPLICATE*" And Not rs_buscaCob.EOF And Not rs_buscaCob.BOF Then
                GoTo Proximo
            ElseIf UCase(lobjCobert.UltimoErro) Like "*DUPLICATE*" And lobjstcCobert.Cod_Ramo = e_CodRamo_Riscos_diversos Then
                GoTo Proximo
            Else
                pMensagem = lobjCobert.UltimoErro
                GoTo SAIDA
            End If
        End If
        If lobjstcCobert.Cod_Cobert = 1 And lValorRisco = 0 Then
            lobjstcPedLoc.Val_Vr = lobjstcCobert.Val_IS
            lobjstcPedLoc.Per_IS_Vr = 100
            If Not lobjPedLoc.mfAlterarValorRisco(lobjstcPedLoc, lPropTabela, "") Then
                pMensagem = lobjPedLoc.UltimoErro
                GoTo SAIDA
            End If
        End If
                
        'Carregar cole��o de nova franquia.
        If M46V999.gPrograma = "D046O003" And lobjstcPedLoc.Tip_Emissao = e_TipEmissao_Endosso Then
            lSql = "Select "
            lSql = lSql & "Cod_Ramo,"
            lSql = lSql & "Cod_Modalidade,"
            lSql = lSql & "Cod_Tarif,"
            lSql = lSql & "Cod_Cobert,"
            lSql = lSql & "Cod_Evento,"
            lSql = lSql & "Num_Ped,"
            lSql = lSql & "Num_Apol,"
            lSql = lSql & "Num_Item,"
            lSql = lSql & "Num_Endosso,"
            lSql = lSql & "NUM_PI,"
            lSql = lSql & "NOSSO_NUMERO,"
            lSql = lSql & "Val_Min_Franq,"
            lSql = lSql & "Per_Franq,"
            lSql = lSql & "Texto_Franquia,"
            lSql = lSql & "Texto_Observacoes,"
            lSql = lSql & "Multiplo_Franquia"
            lSql = lSql & " from ramosdiversos.dbo.tab_ped_cobert_franquia"
            lSql = lSql & " where num_apol = " & prstItem("Num_Apol")
            lSql = lSql & " and cod_cobert = " & rsAPOL!Cod_Cobert
            lSql = lSql & " and num_item = " & rsAPOL!Num_Item
            
            If Not M46V999.gfObtRegistro(pBdP00APOL, lSql, rsNovaFranquia, pMensagem) Then
                GoTo SAIDA
            End If
        
            Set lobjstcMultFraq = New stcA46V738B
            Set lobjMultFraq = New clsA46V738A
            Set lobjMultFraq.ConexaoBD = lobjCobert.ConexaoBD
            Set lobjMultFraq.M46V999 = lobjCobert.M46V999
        
            Do Until rsNovaFranquia.EOF
                lobjstcMultFraq.Cod_Ramo = rsNovaFranquia!Cod_Ramo
                lobjstcMultFraq.Cod_Modalidade = lobjstcCobert.Tab_Cobert
                lobjstcMultFraq.Cod_Tarif = lobjstcPedLoc.Cod_Iden_Tarif
                lobjstcMultFraq.Cod_Cobert = lobjstcCobert.Cod_Cobert
                lobjstcMultFraq.Cod_Evento = rsNovaFranquia!Cod_Evento
                lobjstcMultFraq.Num_Ped = pNumPed
                lobjstcMultFraq.Num_Apol = lobjstcCobert.Num_Apol
                lobjstcMultFraq.Num_Item = lobjstcCobert.Num_Item
                lobjstcMultFraq.Num_Endosso = lobjstcCobert.Num_Endo
                lobjstcMultFraq.NUM_PI = lobjstcCobert.NUM_PI
                lobjstcMultFraq.nosso_numero = lobjstcCobert.nosso_numero
                lobjstcMultFraq.Val_Min_Franq = rsNovaFranquia!Val_Min_Franq
                lobjstcMultFraq.Per_Franq = rsNovaFranquia!Per_Franq
                lobjstcMultFraq.Texto_Franquia = rsNovaFranquia!Texto_Franquia
                lobjstcMultFraq.Texto_Observacoes = rsNovaFranquia!Texto_Observacoes
                lobjstcMultFraq.Multiplo_Franquia = rsNovaFranquia!Multiplo_Franquia
                
                If Not lobjMultFraq.mfIncluir(lobjstcMultFraq, lPropTabela, "", tagDynaTrace) Then
                    pMensagem = lobjMultFraq.UltimoErro
                    GoTo SAIDA
                End If

                rsNovaFranquia.MoveNext
            Loop
    
            Set lobjstcCobert.ColecaoNovaFranquia = lobjMultFraq.Colecao
        End If

Proximo:
        rsAPOL.MoveNext
    Loop


    'ProjetoFormularios
    '
    Dim lNomeTab2 As String


    If M46V999.gBaseLocal = True Then
        lNomeTab2 = ""
    Else
        lNomeTab2 = "RAMOSDIVERSOS.dbo."
    End If

    lSql = " Select * from " & lNomeTab2 & "Tab_Ped_Loc_Pessoas WITH (NOLOCK) Where "
    If pNumCotac_LMI = "" Then
        lSql = lSql & " Num_Apol = " & prstItem("Num_Apol")
        lSql = lSql & " and Num_Endo = " & prstItem("Num_Endo")
    Else
        lSql = lSql & " NOSSO_NUMERO = " & M46V999.gfFormatarCamposSQL(pNumCotac_LMI)
    End If
    lSql = lSql & " and Num_Item = " & prstItem("Num_Item")

    If Not M46V999.gfObtRegistro(pBdP00APOL, lSql, rsAPOL, pMensagem) Then
        GoTo SAIDA
    End If
    
    If rsAPOL.EOF And pNumCotac_LMI = "" Then
        ' N�o encontrou benefici�rios com o n�mero do endosso, procurar ent�o na ap�lice
        lSql = " Select * from " & lNomeTab2 & "Tab_Ped_Loc_Pessoas WITH (NOLOCK) Where "
        lSql = lSql & " Num_Apol = " & prstItem("Num_Apol")
        lSql = lSql & " and Num_Endo = 0"
        lSql = lSql & " and Num_Item = " & prstItem("Num_Item")
    
        If Not M46V999.gfObtRegistro(pBdP00APOL, lSql, rsAPOL, pMensagem) Then
            GoTo SAIDA
        End If
    End If

    If Not rsAPOL.EOF Then
        Dim lobjPessoas As New clsA46V735A
        Dim lobjstcPessoas As New stcA46V735B

        Set lobjPessoas = New clsA46V735A
        Set lobjPessoas.M46V999 = M46V999

        Set lobjPessoas.ConexaoBD = lobjPedLoc.ConexaoBD
        Do While Not rsAPOL.EOF
            Set lobjstcPessoas = New stcA46V735B
            Set lobjstcPessoas.M46V999 = M46V999

            With lobjstcPessoas
                .nosso_numero = lobjstcPedLoc.nosso_numero
                .NUM_PI = lobjstcPedLoc.NUM_PI
                .Num_Item = lobjstcPedLoc.Num_Item
                .Tip_Loc_Pessoa = M46V999.gfTratarNulo(rsAPOL("Tip_Loc_Pessoa"), tpoNumerico)
                .Nom_Pessoa = M46V999.gfTratarNulo(rsAPOL("Nom_Pessoa"), tpoString)
                .Tip_Pessoa = M46V999.gfTratarNulo(rsAPOL("Tip_Pessoa"), tpoNumerico)
                .Num_CPF_CNPJ = M46V999.gfTratarNulo(rsAPOL("Num_CPF_CNPJ"), tpoNumerico)
                .Tip_Cobert = M46V999.gfTratarNulo(rsAPOL("Tip_Cobert"), tpoNumerico)
                .Per_Participacao = M46V999.gfTratarNulo(rsAPOL("Per_Participacao"), tpoNumerico)
            End With
            Call lobjPessoas.mfIncluir(lobjstcPessoas, lPropTabela, "", tagDynaTrace)
            rsAPOL.MoveNext
        Loop
    End If

    '[FRH]->
    Call M46V999.gpFecha2(bdP00MULTI)
    '<-[FRH]
    Set lobjInsp = New clsA46V717A
    Set lobjInsp.M46V999 = M46V999

    Set lObjstcInsp = New stcA46V717B
    lObjstcInsp.nosso_numero = pNosNumero
    lObjstcInsp.Num_Item = lobjstcPedLoc.Num_Item
    lObjstcInsp.Num_Logr_Risco = val(lNumero)
    'corre��o
    lObjstcInsp.Cod_Ramo = lobjstcPedLoc.Cod_Ramo
    lObjstcInsp.Num_Ped = lobjstcPedLoc.Num_Ped
    lObjstcInsp.NUM_PI = lobjstcPedLoc.NUM_PI
    '
    If Not lobjInsp.mfIncluir(lObjstcInsp, lPropTabela, "", tagDynaTrace) Then
        pMensagem = lobjInsp.UltimoErro
        GoTo SAIDA
    End If


    'Ficha 1000409 - Inclus�o de Fatores de Prote��o em Renova��es Yasuda
    If pCod_Dizer = 0 Then
        If M46V999.gBaseLocal = True Then
            lNomeTab = "TAB_APOL"
        Else
            lNomeTab = "RamosDiversos.dbo.Tab_Ped"
        End If
    End If
    '    lSql = " Select * from " & lNomeTab & "_FAT_RISCO "
    '    If (M46V999.gBaseLocal = False And pBdP00APOL Like "*Provider=MSDASQL.*" And M46V999.gPrograma <> "D046O003") Or pNumCotac_LMI <> "" Then
    '        lSql = lSql & " inner join " & IIf(pNumCotac_LMI = "", "GED", "RAMOSDIVERSOS") & ".DBO.TAB_PED  as ctrl on " & lNomeTab & "_FAT_RISCO.num_apol = ctrl.num_apol"
    '        lSql = lSql & " and ctrl.num_endo = " & lNomeTab & "_FAT_RISCO.num_endo"
    '        lSql = lSql & " and ctrl.num_PI = " & lNomeTab & "_FAT_RISCO.num_PI"
    '        If pNumCotac_LMI <> "" Then '1100177 - Renova��o Apolice LMI
    '            lSql = lSql & " and ctrl.NOSSO_NUMERO = " & lNomeTab & "_FAT_RISCO.NOSSO_NUMERO"
    '        End If
    '    End If
    '    If pNumCotac_LMI <> "" Then '1100177 - Renova��o Apolice LMI
    '        lSql = lSql & " Where " & lNomeTab & "_FAT_RISCO.NOSSO_NUMERO = " & M46V999.gfFormatarCamposSQL(pNumCotac_LMI)
    '        lSql = lSql & " and ctrl.NOSSO_NUMERO = " & lNomeTab & "_FAT_RISCO.NOSSO_NUMERO"
    '    Else
    '        lSql = lSql & " Where " & lNomeTab & "_FAT_RISCO.Num_Apol = " & prstItem("Num_Apol")
    '        lSql = lSql & " and " & lNomeTab & "_FAT_RISCO.num_Endo = " & prstItem("NUm_Endo2")
    '        'CR13986 - OS PROTECIONAIS N�O ESTAVAM RETORNANDO QUANDO DA CRIA��O DE COTA��O DE ENDOSSO.
    '        'lSql = lSql & " and " & lNomeTab & "_FAT_RISCO.num_ped = " & lObjstcPedLoc.Num_Ped
    '    End If
    '    lSql = lSql & " and " & lNomeTab & "_FAT_RISCO.num_item = " & prstItem("Num_Item")

    lSql = " Select DISTINCT FAT_RISCO.COD_RAMO,FAT_RISCO.COD_FATOR,FAT_RISCO.DSC_COMPL_FATOR from " & lNomeTab & "_FAT_RISCO as FAT_RISCO WITH (NOLOCK)"
    If (M46V999.gBaseLocal = False And pBdP00APOL Like "*Provider=SQLOLEDB.*" And M46V999.gPrograma <> "D046O003") Or pNumCotac_LMI <> "" Then
        lSql = lSql & " inner join " & IIf(pNumCotac_LMI = "", "GED", "RAMOSDIVERSOS") & ".DBO.TAB_PED  as ctrl WITH (NOLOCK) on FAT_RISCO.num_apol = ctrl.num_apol"
        lSql = lSql & " and ctrl.num_endo = FAT_RISCO.num_endo"
        lSql = lSql & " and ctrl.num_PI = FAT_RISCO.num_PI"
        If pNumCotac_LMI <> "" Then    '1100177 - Renova��o Apolice LMI
            lSql = lSql & " and ctrl.NOSSO_NUMERO = FAT_RISCO.NOSSO_NUMERO"
        End If
    End If
    If pNumCotac_LMI <> "" Then    '1100177 - Renova��o Apolice LMI
        lSql = lSql & " Where FAT_RISCO.NOSSO_NUMERO = " & M46V999.gfFormatarCamposSQL(pNumCotac_LMI)
        lSql = lSql & " and ctrl.NOSSO_NUMERO = FAT_RISCO.NOSSO_NUMERO"
    Else
        lSql = lSql & " Where FAT_RISCO.Num_Apol = " & prstItem("Num_Apol")
        lSql = lSql & " and FAT_RISCO.num_Endo = " & prstItem("NUm_Endo2")
        'CR13986 - OS PROTECIONAIS N�O ESTAVAM RETORNANDO QUANDO DA CRIA��O DE COTA��O DE ENDOSSO.
        'lSql = lSql & " and " & lNomeTab & "_FAT_RISCO.num_ped = " & lObjstcPedLoc.Num_Ped
    End If
    lSql = lSql & " and FAT_RISCO.num_item = " & prstItem("Num_Item")
    On Error Resume Next
    If Not M46V999.gfObtRegistro(pBdP00APOL, lSql, rsAPOL, pMensagem) Then
        GoTo ControleErro
    End If

    Set lObjPedFator = New clsA46V720A
    Set lObjPedFator.M46V999 = M46V999

    Set lObjPedFator.ConexaoBD = lobjPedLoc.ConexaoBD
    Do While Not rsAPOL.EOF
        If rsAPOL("Cod_Ramo") = 113 And rsAPOL("Cod_Fator") = 7 Then
            'Ignorar- nao incluir
        Else
            Set lobjstcPedFator = New stcA46V720B
            With lobjstcPedFator
                .Cod_Ramo = M46V999.gfTratarNulo(rsAPOL("Cod_Ramo"), tpoNumerico)
                .nosso_numero = pNosNumero
                .Num_Ped = pNumPed
                If M46V111.gNum_Pi = 0 Then
                    .NUM_PI = Right(Year(Date), 2) & Format(pNumPed, "00000000")
                Else
                    .NUM_PI = M46V111.gNum_Pi
                End If
                .Num_Item = lobjstcPedLoc.Num_Item
                .Cod_Fator = M46V999.gfTratarNulo(rsAPOL("Cod_Fator"), tpoNumerico)
                .Dsc_Compl_Fator = M46V999.gfTratarNulo(rsAPOL("DSC_COMPL_FATOR"), tpoNumerico)
            End With
            If Not lObjPedFator.mfIncluir(lobjstcPedFator, lPropTabela, "", tagDynaTrace) Then
                pMensagem = lObjPedFator.UltimoErro
                GoTo SAIDA
            End If
        End If
        rsAPOL.MoveNext
    Loop
    On Error GoTo ControleErro

    'gravar fator de prote��o
    '1200539 - Tarifa 1038 e 1039 Junho 2012 - Item 8
    'Retirado 1200733
    If lobjstcPedLoc.Cod_Ramo = e_CodRamo_Residencial And _
       lobjstcPedLoc.gLiberaDescontoEspecial Then

        Set lobjstcPedFator = New stcA46V720B
        Set lObjPedFator = New clsA46V720A
        Set lObjPedFator.M46V999 = M46V999

        With lobjstcPedFator
            .nosso_numero = lobjstcPedLoc.nosso_numero
            .Num_Ped = lobjstcPedLoc.Num_Ped
            .Cod_Ramo = lobjstcPedLoc.Cod_Ramo
            .Num_Item = lobjstcPedLoc.Num_Item
            .Cod_Org = lobjstcPedLoc.Cod_Org
            .Cod_Fator = 7
            .Desc_Fator = "DESCONTO ESPECIAL"
            Set .ObjPai = lobjstcPedLoc
        End With

        Call lObjPedFator.mfIncluir(lobjstcPedFator, lPropTabela, "", tagDynaTrace)    'Ignora Erros!!!

    End If
        
    'Grava os bens
        lSql = "Select "
        lSql = lSql & "NOSSO_NUMERO,"
        lSql = lSql & "Cod_Org,"
        lSql = lSql & "Cod_Ramo,"
        lSql = lSql & "Num_Ped,"
        lSql = lSql & "Num_Apol,"
        lSql = lSql & "Num_Endo,"
        lSql = lSql & "Num_Item,"
        lSql = lSql & "Dsc_Bem,"
        lSql = lSql & "Val_Bem,"
        lSql = lSql & "Cod_Bem,"
        lSql = lSql & "Cod_Seq"
        lSql = lSql & " from ramosdiversos.dbo.Tab_ped_Loc_Bem"
        lSql = lSql & " where num_apol = " & prstItem("Num_Apol")
        lSql = lSql & " and num_item = " & prstItem("Num_item")
        lSql = lSql & " and Num_Endo = " & prstItem("Num_Endo")

        
        If Not M46V999.gfObtRegistro(pBdP00APOL, lSql, rsAPOL, pMensagem) Then
            GoTo SAIDA
        End If
    
        Set lobjstcPedLocBem = New stcA46V715B
        Set lobjPedLocBem = New clsA46V715A
        Set lobjPedLocBem.ConexaoBD = lobjPedLoc.ConexaoBD
        Set lobjPedLocBem.M46V999 = lobjPedLoc.M46V999
        
        Do Until rsAPOL.EOF
            lobjstcPedLocBem.nosso_numero = lobjstcPedLoc.nosso_numero
            lobjstcPedLocBem.Cod_Org = lobjstcPedLoc.Cod_Org
            lobjstcPedLocBem.Cod_Ramo = lobjstcPedLoc.Cod_Ramo
            lobjstcPedLocBem.Num_Ped = lobjstcPedLoc.Num_Ped
            lobjstcPedLocBem.Num_Apol = lobjstcPedLoc.Num_Apol
            lobjstcPedLocBem.Num_Endo = lobjstcPedLoc.Num_Endo
            lobjstcPedLocBem.Num_Item = lobjstcPedLoc.Num_Item
            lobjstcPedLocBem.Dsc_Bem = rsAPOL!Dsc_Bem
            lobjstcPedLocBem.Val_Bem = rsAPOL!Val_Bem
            lobjstcPedLocBem.Cod_Bem = rsAPOL!Cod_Bem
            lobjstcPedLocBem.Cod_Seq = rsAPOL!Cod_Seq
            
            If Not lobjPedLocBem.mfIncluir(lobjstcPedLocBem, lPropTabela, "", tagDynaTrace) Then
                pMensagem = lobjPedLocBem.UltimoErro
                GoTo SAIDA
            End If
            
            rsAPOL.MoveNext
        Loop

        Set lobjstcPedLoc.ColecaoPedLocBem = lobjPedLocBem.Colecao
        
        'Grava bicicletas
        lSql = "Select "
        lSql = lSql & "NOSSO_NUMERO,"
        lSql = lSql & "Cod_Org,"
        lSql = lSql & "Cod_Ramo,"
        lSql = lSql & "Num_Ped,"
        lSql = lSql & "Num_Apol,"
        lSql = lSql & "Num_Endo,"
        lSql = lSql & "Num_Item,"
        lSql = lSql & "Num_PI,"
        lSql = lSql & "Descricao,"
        lSql = lSql & "Valor,"
        lSql = lSql & "Cod_Seq"
        lSql = lSql & " from ramosdiversos.dbo.Tab_ped_Loc_Bicicleta"
        lSql = lSql & " where num_apol = " & prstItem("Num_Apol")
        lSql = lSql & " and num_item = " & prstItem("Num_item")
        lSql = lSql & " and Num_Endo = " & prstItem("Num_Endo")

        If Not M46V999.gfObtRegistro(pBdP00APOL, lSql, rsAPOL, pMensagem) Then
            GoTo SAIDA
        End If
                
        Set lobjstcPedLocBicicleta = New stcA46V744B
        Set lobjPedLocBicicleta = New clsA46V744A
        Set lobjPedLocBicicleta.ConexaoBD = lobjPedLoc.ConexaoBD
        Set lobjPedLocBicicleta.M46V999 = lobjPedLoc.M46V999
        
        Do Until rsAPOL.EOF
            lobjstcPedLocBicicleta.nosso_numero = lobjstcPedLoc.nosso_numero
            lobjstcPedLocBicicleta.Cod_Org = lobjstcPedLoc.Cod_Org
            lobjstcPedLocBicicleta.Cod_Ramo = lobjstcPedLoc.Cod_Ramo
            lobjstcPedLocBicicleta.Num_Ped = lobjstcPedLoc.Num_Ped
            lobjstcPedLocBicicleta.Num_Apol = lobjstcPedLoc.Num_Apol
            lobjstcPedLocBicicleta.Num_Endo = lobjstcPedLoc.Num_Endo
            lobjstcPedLocBicicleta.Num_Item = lobjstcPedLoc.Num_Item
            lobjstcPedLocBicicleta.NUM_PI = lobjstcPedLoc.NUM_PI
            lobjstcPedLocBicicleta.Descricao = rsAPOL!Descricao
            lobjstcPedLocBicicleta.Valor = rsAPOL!Valor
            lobjstcPedLocBicicleta.Cod_Seq = rsAPOL!Cod_Seq
            
            If Not lobjPedLocBicicleta.mfIncluir(lobjstcPedLocBicicleta, lPropTabela, "", tagDynaTrace) Then
                pMensagem = lobjPedLocBicicleta.UltimoErro
                GoTo SAIDA
            End If
            
            rsAPOL.MoveNext
        Loop

        Set lobjstcPedLoc.ColecaoPedLocBicicleta = lobjPedLocBicicleta.Colecao

        
        'Grava AllRisks
        lSql = "Select "
        lSql = lSql & "NOSSO_NUMERO,"
        lSql = lSql & "Cod_Org,"
        lSql = lSql & "Cod_Ramo,"
        lSql = lSql & "Num_Ped,"
        lSql = lSql & "Num_Apol,"
        lSql = lSql & "Num_Endo,"
        lSql = lSql & "Num_Item,"
        lSql = lSql & "Num_PI,"
        lSql = lSql & "Descricao,"
        lSql = lSql & "Valor,"
        lSql = lSql & "Cod_Seq"
        lSql = lSql & " from ramosdiversos.dbo.Tab_Ped_All_Risks"
        lSql = lSql & " where num_apol = " & prstItem("Num_Apol")
        lSql = lSql & " and num_item = " & prstItem("Num_item")
        lSql = lSql & " and Num_Endo = " & prstItem("Num_Endo")

        If Not M46V999.gfObtRegistro(pBdP00APOL, lSql, rsAPOL, pMensagem) Then
            GoTo SAIDA
        End If
                
        Set lobjstcPedAllRisks = New stcA46V747B
        Set lobjPedAllRisks = New clsA46V747A
        Set lobjPedAllRisks.ConexaoBD = lobjPedLoc.ConexaoBD
        Set lobjPedAllRisks.M46V999 = lobjPedLoc.M46V999
        
        Do Until rsAPOL.EOF
            lobjstcPedAllRisks.nosso_numero = lobjstcPedLoc.nosso_numero
            lobjstcPedAllRisks.Cod_Org = lobjstcPedLoc.Cod_Org
            lobjstcPedAllRisks.Cod_Ramo = lobjstcPedLoc.Cod_Ramo
            lobjstcPedAllRisks.Num_Ped = lobjstcPedLoc.Num_Ped
            lobjstcPedAllRisks.Num_Apol = lobjstcPedLoc.Num_Apol
            lobjstcPedAllRisks.Num_Endo = lobjstcPedLoc.Num_Endo
            lobjstcPedAllRisks.Num_Item = lobjstcPedLoc.Num_Item
            lobjstcPedAllRisks.NUM_PI = lobjstcPedLoc.NUM_PI
            lobjstcPedAllRisks.Descricao = rsAPOL!Descricao
            lobjstcPedAllRisks.Valor = rsAPOL!Valor
            lobjstcPedAllRisks.Cod_Seq = rsAPOL!Cod_Seq
            
            If Not lobjPedAllRisks.mfIncluir(lobjstcPedAllRisks, lPropTabela, "", tagDynaTrace) Then
                pMensagem = lobjPedAllRisks.UltimoErro
                GoTo SAIDA
            End If
            
            rsAPOL.MoveNext
        Loop

        Set lobjstcPedLoc.ColecaoAllRisks = lobjPedAllRisks.Colecao
        
        
        'Grava Obras Arte
        lSql = "Select "
        lSql = lSql & "NOSSO_NUMERO,"
        lSql = lSql & "Cod_Org,"
        lSql = lSql & "Cod_Ramo,"
        lSql = lSql & "Num_Ped,"
        lSql = lSql & "Num_Apol,"
        lSql = lSql & "Num_Endo,"
        lSql = lSql & "Num_Item,"
        lSql = lSql & "Num_PI,"
        lSql = lSql & "Descricao,"
        lSql = lSql & "Valor,"
        lSql = lSql & "Cod_Seq"
        lSql = lSql & " from ramosdiversos.dbo.Tab_Ped_Obras_Arte"
        lSql = lSql & " where num_apol = " & prstItem("Num_Apol")
        lSql = lSql & " and num_item = " & prstItem("Num_item")
        lSql = lSql & " and Num_Endo = " & prstItem("Num_Endo")

        If Not M46V999.gfObtRegistro(pBdP00APOL, lSql, rsAPOL, pMensagem) Then
            GoTo SAIDA
        End If
                
        Set lobjstcPedObrasArte = New stcA46V748B
        Set lobjPedObrasArte = New clsA46V748A
        Set lobjPedObrasArte.ConexaoBD = lobjPedLoc.ConexaoBD
        Set lobjPedObrasArte.M46V999 = lobjPedLoc.M46V999
        
        Do Until rsAPOL.EOF
            lobjstcPedObrasArte.nosso_numero = lobjstcPedLoc.nosso_numero
            lobjstcPedObrasArte.Cod_Org = lobjstcPedLoc.Cod_Org
            lobjstcPedObrasArte.Cod_Ramo = lobjstcPedLoc.Cod_Ramo
            lobjstcPedObrasArte.Num_Ped = lobjstcPedLoc.Num_Ped
            lobjstcPedObrasArte.Num_Apol = lobjstcPedLoc.Num_Apol
            lobjstcPedObrasArte.Num_Endo = lobjstcPedLoc.Num_Endo
            lobjstcPedObrasArte.Num_Item = lobjstcPedLoc.Num_Item
            lobjstcPedObrasArte.NUM_PI = lobjstcPedLoc.NUM_PI
            lobjstcPedObrasArte.Descricao = rsAPOL!Descricao
            lobjstcPedObrasArte.Valor = rsAPOL!Valor
            lobjstcPedObrasArte.Cod_Seq = rsAPOL!Cod_Seq
            
            If Not lobjPedObrasArte.mfIncluir(lobjstcPedObrasArte, lPropTabela, "", tagDynaTrace) Then
                pMensagem = lobjPedObrasArte.UltimoErro
                GoTo SAIDA
            End If
            
            rsAPOL.MoveNext
        Loop

        Set lobjstcPedLoc.ColecaoObrasArte = lobjPedObrasArte.Colecao

    '*********************** FInalizado com Sucesso
    mfGraRenItem = True
    '***********************
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)

SAIDA:

    GoSub FechaObj
    Exit Function

ControleErro:
    If Err.Description <> "" Then
        Call M46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    End If
    Call M46V999.gpGraLog(0, "gfGraRenItem ERRO - " & Err & " " & Err.Description & " " & pMensagem)
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "gfGraRenItem ERRO - " & Err & " " & Err.Description & " " & pMensagem, 2, Class_Name, methodName)
    GoTo SAIDA
FechaObj:
    Call M46V999.gpFecha2(bdP00MULTI)
    Set rsAPOL = Nothing
    Set rs_buscaCob = Nothing
    Set lobjPedLoc = Nothing
    Set lobjstcPedLoc = Nothing
    Set lObjClau = Nothing
    Set lobjstcClau = Nothing
    Set lobjCobert = Nothing
    Set lobjstcCobert = Nothing
    Set lobjInsp = Nothing
    Set lObjstcInsp = Nothing
    Set lObjPedFator = Nothing
    Set lobjstcPedFator = Nothing
    Set lrsFranquia = Nothing
    Set lobjMultFraq = Nothing
    Set lobjstcMultFraq = Nothing
    Set lobjPedLocBem = Nothing
    Set lobjstcPedLocBem = Nothing
    
    Return
End Function




'11/01/2010 - Ficha: 1000003 - franquia e valor de franquia
Private Sub mpBuscaLimFranquiaDesativado(ByRef pobjstcPedLoc As stcA46V708B, ByRef pobjStcCobert As stcA46V704B)
'Desativado Simony
    Dim BDP00Tarifa As ADODB.Connection
    Dim rstTabClas As clsYasRecordSet
    Dim lbdTbm_Fat_Franquia As ADODB.Connection         'Arquivo: P00Multi.
    Dim lMensagem As String         'Mensagem.
    Dim lrsTbm_Fat_Franquia As clsYasRecordSet          'Registro : Tbm_Fator_Franquia.
    Dim lCodSequencia As Integer
    Dim lCEP As String
    Dim lTip_Atividade As Integer     'Utilizado para selecao de Franquia
    Dim lSelect As String

    pobjStcCobert.Per_Franq = 0
    pobjStcCobert.Val_Franq_Min = 0
    pobjStcCobert.Val_Fator_Franq = 0


    If Trim(pobjstcPedLoc.Cod_Clasf) <> "" And Trim(pobjstcPedLoc.Cod_Clasf) <> "0" Then
        '1. Abre arquivo.
        If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, BDP00Tarifa, lMensagem) = False Then
            Call M46V999.gpGraLog(1, lMensagem)
            Exit Sub
        End If

        '3. Obt�m produto sem t�rmino de vig�ncia ou com t�rmino de vig�ncia superior � data do dia.
        lSelect = " Select distinct dsc_clasf, COD_CLASF, " & _
                  "COD_COMPL_CLASF , tab_Clasf, Tip_Atividade  from Tbm_Classificacao WITH (NOLOCK) " & _
                " Where cod_Ramo  = " & pobjstcPedLoc.Cod_Ramo & _
                " and cod_Clasf = '" & pobjstcPedLoc.Cod_Clasf & "'" & _
                " and Cod_Compl_Clasf = '" & pobjstcPedLoc.Cod_Compl_Clasf & "'" & _
                " and Cod_Tarif = " & pobjstcPedLoc.Cod_Iden_Tarif & _
                " and Cod_Sequencia <= " & M46V999.gCodSequencia
        lSelect = lSelect & " and  tab_clasf = " & M46V111.gCod_plano

        If M46V999.gfObtRegistro(BDP00Tarifa, lSelect, rstTabClas, lMensagem) = False Then
            Call M46V999.gpFecha2(BDP00Tarifa)
            Call M46V999.gpGraLog(1, lMensagem)
            Exit Sub
        End If

        If Not rstTabClas.EOF Then
            lTip_Atividade = rstTabClas("Tip_Atividade")
        End If
    End If




    If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, lbdTbm_Fat_Franquia, lMensagem) = False Then
        Call M46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If

    lCodSequencia = M46V999.gCodSequencia

NovaSequencia1:

    '1 - Procura Registro de exce��o com a chave completa
    'Procura o valor do Val_Franq_Min e Perc_Franq
    lCEP = val(Trim(Replace(pobjstcPedLoc.Num_Cep_Risco, "-", "")))

    lSelect = " Select Val_Franq_Min as Min , Perc_Franquia as Perc, Fator_Ajuste as Fator from Tbm_Fator_Franquia WITH (NOLOCK) "
    lSelect = lSelect & " Where Cod_Ramo = " & pobjstcPedLoc.Cod_Ramo
    lSelect = lSelect & " and Cod_Modalidade = " & M46V111.gCod_plano
    lSelect = lSelect & " and Cod_Tip_Reg = 'E' "
    If M46V999.gCodProduto >= 940 Then
        lSelect = lSelect & " and ((Cod_Clasf = '" & pobjstcPedLoc.Cod_Clasf & "'"
        lSelect = lSelect & " and Cod_Compl_Clasf = '" & pobjstcPedLoc.Cod_Compl_Clasf & "')"
        lSelect = lSelect & " or (Cod_Clasf = '0'))"
    Else
        lSelect = lSelect & " and Cod_Clasf = '" & pobjstcPedLoc.Cod_Clasf & "'"
        lSelect = lSelect & " and Cod_Compl_Clasf = '" & pobjstcPedLoc.Cod_Compl_Clasf & "'"
    End If
    lSelect = lSelect & " and Cod_Cobert = " & pobjStcCobert.Cod_Cobert
    lSelect = lSelect & " and Cod_Construcao in (0 , " & pobjstcPedLoc.Cod_Constr & ")"
    lSelect = lSelect & " and Cod_Tarif = " & pobjstcPedLoc.Cod_Iden_Tarif
    lSelect = lSelect & " and Cod_Sequencia <= " & lCodSequencia

    If M46V999.gCodProduto >= 940 Then
        lSelect = lSelect & " AND ((CEP_INICIO <= " & lCEP
        lSelect = lSelect & " AND CEP_FIM >= " & lCEP & " ) "
        lSelect = lSelect & " OR (CEP_INICIO =  0 ))"
        lSelect = lSelect & " order by Prioridade desc "
    End If

    If M46V999.gfObtRegistro(lbdTbm_Fat_Franquia, lSelect, lrsTbm_Fat_Franquia, lMensagem) = False Then
        Call M46V999.gpFecha2(lbdTbm_Fat_Franquia)
        Call M46V999.gpGraLog(2, lMensagem)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    '2 - Se n�o encontrar a chave completa vai procurar o padr�o
    If lrsTbm_Fat_Franquia.EOF Then
        If lCodSequencia > 0 Then
            lCodSequencia = lCodSequencia - 1

            GoTo NovaSequencia1
        Else
            GoTo ProcuraDefault
        End If
    End If

    If lrsTbm_Fat_Franquia("Min") = 0 Then
        GoTo ProcuraDefault
    End If

    pobjStcCobert.Per_Franq = lrsTbm_Fat_Franquia("Perc")
    pobjStcCobert.Val_Franq_Min = lrsTbm_Fat_Franquia("Min")
    pobjStcCobert.Val_Fator_Franq = lrsTbm_Fat_Franquia("Fator")

    GoTo FimBusca

ProcuraDefault:



    '3 - Procura o registro somente com Ramo, Modalidade, Cobertura e TipoAtividade
    'Procura o valor do Val_Franq_Min, Perc_Franq

    lCodSequencia = M46V999.gCodSequencia

NovaSequencia2:

    lSelect = " Select Val_Franq_Min as Min , Perc_Franquia as Perc, Fator_Ajuste as Fator  from Tbm_Fator_Franquia WITH (NOLOCK) "
    lSelect = lSelect & " Where Cod_Ramo = " & pobjstcPedLoc.Cod_Ramo
    lSelect = lSelect & " and Cod_Modalidade = " & M46V111.gCod_plano
    lSelect = lSelect & " and Cod_Tip_Reg = 'D' "
    lSelect = lSelect & " and Cod_Cobert = " & pobjStcCobert.Cod_Cobert
    lSelect = lSelect & " and Tip_Atividade in (3," & lTip_Atividade & ")"
    lSelect = lSelect & " and Cod_Tarif = " & pobjstcPedLoc.Cod_Iden_Tarif
    lSelect = lSelect & " and Cod_Sequencia <= " & lCodSequencia

    If M46V999.gfObtRegistro(lbdTbm_Fat_Franquia, lSelect, lrsTbm_Fat_Franquia, lMensagem) = False Then
        Call M46V999.gpFecha2(lbdTbm_Fat_Franquia)
        Call M46V999.gpGraLog(2, lMensagem)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    If lrsTbm_Fat_Franquia.EOF Then
        If lCodSequencia > 0 Then
            lCodSequencia = lCodSequencia - 1

            GoTo NovaSequencia2
        Else
            GoTo FimBusca
        End If
    End If

    If lrsTbm_Fat_Franquia("Min") = 0 Then
        GoTo FimBusca
    End If

    pobjStcCobert.Per_Franq = lrsTbm_Fat_Franquia("Perc")
    pobjStcCobert.Val_Franq_Min = lrsTbm_Fat_Franquia("Min")
    pobjStcCobert.Val_Fator_Franq = lrsTbm_Fat_Franquia("Fator")

FimBusca:
    Call M46V999.gpFecha2(lbdTbm_Fat_Franquia)

    '1001135 - Corretor Volks - Franquia Especial Concessionaria
    If M46V999.gCodProduto >= 1002 Then
        If (InStr(1, "00008001/00008002/00008003/00008004/00008005/00008006/00008008/00008011/00008014/00008028", Format(M46V999.gCod_Corretor, "00000000"), vbTextCompare) > 0) And _
           (val(pobjstcPedLoc.Cod_Clasf) = M46V999.gfBuscarDePara(22, 2) And val(pobjstcPedLoc.Cod_Compl_Clasf) = M46V999.gfBuscarDePara(22, 2, 13)) And _
           (InStr(1, "0038/0064/0066", Format(pobjStcCobert.Cod_Cobert, "0000"), vbTextCompare) > 0) Then
            pobjStcCobert.Per_Franq = 15
            pobjStcCobert.Val_Franq_Min = 2000
        End If
    End If
    Set BDP00Tarifa = Nothing
    Set rstTabClas = Nothing
    Set lbdTbm_Fat_Franquia = Nothing
    Set lrsTbm_Fat_Franquia = Nothing



End Sub

Public Function gfImportarTransmissao(pNossoNumero As String, _
                                      pCodUnicidade As Long, _
                                      pNovNosNumero As String, _
                                      pNovNumPedido As String, _
                                      pNumeroPI As Long, _
                                      pDatRecebimento As Long, _
                                      ByRef pMensagem As String) As Boolean


'Fun��o: Rotina que transfere os registros das Tabelas de FTP para base quente
'Par�metros de entrada...pAntNumItem....Anterior: n�mero do item.
'                        pmfNovNumItem....Novo: n�mero do item.
'Par�metro de sa�da......pMensagem......Mensagem.

    Dim bdBasDados As ADODB.Connection    'Base de dados.
    Dim flTabela As clsYasField   'Campos da tabela.
    Dim rsTabela As clsYasRecordSet    'Registro da tabela.
    Dim lCount As Long        'Utilizado no for...next.
    Dim lIdeTabela As String          'Nome da tabela.
    Dim lContinuar As Boolean         'Indicador se continua o procedimento ou n�o.
    Dim Dummy As Variant
    Dim lSelect As String
    Dim lInsert As String

    On Error GoTo ControleErro

    gfImportarTransmissao = False

    '1. Obt�m registros da tabela.
    If M46V999.gfAbrBasDados("P00MPED", M46V999.gTipP00MPED, M46V999.gEndP00MPED, bdBasDados, _
                             pMensagem) = False Then Exit Function

    For lCount = 1 To 31
        Select Case lCount
        Case 1
            lIdeTabela = "TAB_PED"
        Case 2
            lIdeTabela = "TAB_PED_CORR"
        Case 3
            lIdeTabela = "TAB_PED_COSS"
        Case 4
            lIdeTabela = "TAB_PED_JUROS"
        Case 5
            lIdeTabela = "TAB_PED_VENC"
        Case 6
            lIdeTabela = "TAB_PED_ENDER"
        Case 7
            lIdeTabela = "TAB_CTRL_EMIS"
        Case 8
            lIdeTabela = "TAB_PED_LOC"
        Case 9
            lIdeTabela = "TAB_PED_CLAU"
        Case 10
            lIdeTabela = "TAB_PED_COBERT"
        Case 11
            lIdeTabela = "TAB_PED_LOC_BEM"
        Case 12
            lIdeTabela = "TAB_PED_LOC_SEGURO"
        Case 13
            lIdeTabela = "TAB_PED_OBS"
        Case 14
            lIdeTabela = "TAB_PED_INSP"
        Case 15
            lIdeTabela = "TAB_PED_AUTO"
        Case 16
            lIdeTabela = "TAB_PED_COBERT_COMPL"
        Case 17
            lIdeTabela = "TAB_PED_FAT_RISCO"
        Case 18
            lIdeTabela = "TAB_PED_ESPEC"
        Case 19
            lIdeTabela = "Tab_Ped_Espec2"
        Case 20
            lIdeTabela = "Tab_Ped_Dizer"
        Case 21
            lIdeTabela = "Tab_Ped_Loc_LMI"
        Case 22
            lIdeTabela = "Tab_Canal_Emiss"
        Case 23
            lIdeTabela = "TAB_CTRL_EMIS_COMPL"
        Case 24
            lIdeTabela = "TAB_PED_LOC_PESSOAS"
        Case 25
            lIdeTabela = "TAB_PED_TELEFONES"
        Case 26
            lIdeTabela = "TAB_PED_LOC_COMPL"
        Case 27
            lIdeTabela = "TAB_PED_COBERT_FRANQUIA"
        Case 28
            lIdeTabela = "TAB_PED_TEXTOS"
        Case 29
            lIdeTabela = "TAB_PED_LIBERACAO_SERASA"
        Case 30
            lIdeTabela = "TAB_PED_ALL_RISKS"
        Case 31
            lIdeTabela = "TAB_PED_OBRAS_ARTE"
        End Select
        lSelect = "SELECT * FROM FTP_" & lIdeTabela & " WITH (NOLOCK) WHERE Nosso_Numero = '" & pNossoNumero & "'"
        If lIdeTabela <> "TAB_PED_TEXTOS" And lIdeTabela <> "TAB_PED_LIBERACAO_SERASA" Then
            lSelect = lSelect & " AND COD_UNICIDADE = " & pCodUnicidade
        End If
        If M46V999.gfObtRegistro(bdBasDados, lSelect, rsTabela, pMensagem) = False Then
            Call M46V999.gpFecha2(bdBasDados)
            Exit Function
        End If
        lContinuar = True
        If rsTabela.EOF = True Then
            lContinuar = False
            If lIdeTabela = "TAB_PED" Then
                pMensagem = "FMP0001 - " & M46V999.gPrefixo & "Pedido n�o localizado (" & pNossoNumero & _
                          " - c�pia da cota��o/proposta.)."
                Exit Function
            End If
        Else
            If lIdeTabela = "TAB_PED" Then
                If rsTabela("Cod_Tip_Doc") = 9 Then
                    If MsgBox("FMP0002 - " & M46V999.gPrefixo & " - Pedido j� importado! Deseja importar novamente?", vbYesNo) = vbNo Then
                        gfImportarTransmissao = True
                        Exit Function
                    End If
                End If
                'Abrindo uma Transa��o
                If M46V999.gfExeSQL(bdBasDados, "Begin Transaction ", pMensagem) <> 0 Then
                    Call M46V999.gpFecha2(bdBasDados)
                    Exit Function
                End If
            End If
        End If
        If lContinuar = True Then
            'Vai atualizar o Cod_tip_Doc para marcar que j� foi carregado
            If lIdeTabela = "TAB_PED" Then
                lSelect = "Update FTP_" & lIdeTabela & " Set Cod_Tip_Doc = 9  WHERE Nosso_Numero = '" & pNossoNumero & "'"
                lSelect = lSelect & " AND COD_UNICIDADE = " & pCodUnicidade
                If M46V999.gfExeSQL(bdBasDados, lSelect, pMensagem) <> 0 Then
                    GoTo ErrRollBack
                    Exit Function
                End If
            End If
            '2. Grava nova cota��o/proposta.
            While Not rsTabela.EOF
                lInsert = "INSERT INTO " & lIdeTabela & " ("
                For Each flTabela In rsTabela.Fields
                    Select Case UCase(flTabela.Name)
                    Case "VAL_RENDAFATU", "TELEFONE_ID", "PED_LOC_PESSOA_ID", "ID_FRANQUIA_AVANCADA"
                    Case "TIP_TRANSMISSAO", "DTHRRECEBTRANSMIS"
                        If lIdeTabela <> "TAB_PED" Then
                            lInsert = lInsert & flTabela.Name & ", "
                        End If
                    Case "COD_UNICIDADE"
                        If lIdeTabela = "TAB_PED" Then
                            lInsert = lInsert & flTabela.Name & ", "
                        End If
                    Case Else
                        lInsert = lInsert & flTabela.Name & ", "
                    End Select
                Next flTabela
                lInsert = Left$(lInsert, Len(lInsert) - 2) & ") VALUES ("
                For Each flTabela In rsTabela.Fields
                    Select Case UCase(flTabela.Name)
                    Case "VAL_RENDAFATU", "TELEFONE_ID", "PED_LOC_PESSOA_ID", "ID_FRANQUIA_AVANCADA"
                    Case "TIP_TRANSMISSAO"
                        If lIdeTabela <> "TAB_PED" Then
                            If IsNull(flTabela.Value) Then
                                lInsert = lInsert & "0, "
                            Else
                                lInsert = lInsert & flTabela.Value & ", "
                            End If
                        End If
                    Case "DTHRRECEBTRANSMIS"
                        If lIdeTabela <> "TAB_PED" Then
                            If M46V999.gfPreenchido(flTabela.Value) Then
                                lInsert = lInsert & "'" & M46V999.gfForAspas(flTabela.Value) & "', "
                            Else
                                lInsert = lInsert & "' ', "
                            End If
                        End If
                    Case "NUM_PI"
                        lInsert = lInsert & pNumeroPI & ", "                             'N�mero do pedido.
                    Case "NOSSO_NUMERO"
                        lInsert = lInsert & "'" & pNovNosNumero & "', "                   'Novo Nosso N�mero.
                    Case "NUM_PED", "NUM_PROP"
                        lInsert = lInsert & pNovNumPedido & ", "                       'Novo Nosso N�mero.
                    Case "SIT_DOCUMENTO"
                        lInsert = lInsert & "1, "       'DOCUMENTO N�O CALCULADO.       'Sit_Documento.
                    Case "DAT_LIBER", "DAT_PRI_CALCULO"
                        lInsert = lInsert & Left(M46V999.gfForData(3), 8) & ", "                        'Dat_Alt.
                    Case "COD_USER_LIBER"
                        lInsert = lInsert & "'COPIA',"                                  'Cod_User_Alt.
                    Case "DAT_DIGIT"
                        lInsert = lInsert & Left(M46V999.gfForData(3), 8) & ","
                    Case "DAT_REC_CIA"
                        lInsert = lInsert & pDatRecebimento & ","                                  'Cod_User_Alt.
                    Case "COD_UNICIDADE"
                        If lIdeTabela = "TAB_PED" Then
                            If IsNull(flTabela.Value) Then
                                lInsert = lInsert & "0, "
                            Else
                                lInsert = lInsert & flTabela.Value & ", "
                            End If
                        End If
                    Case Else
                        Select Case flTabela.Type_
                        Case adBigInt, adBinary, adInteger, adLongVarBinary, adSmallInt, _
                             adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, _
                             adUnsignedTinyInt                          'Num�rico - inteiro.
                            If IsNull(flTabela.Value) Then
                                lInsert = lInsert & "0, "
                            Else
                                lInsert = lInsert & flTabela.Value & ", "
                            End If
                        Case adCurrency, adSingle, adDouble, adNumeric            'Num�rico - com v�rgula.
                            If IsNull(flTabela.Value) Then
                                lInsert = lInsert & "0, "
                            Else
                                lInsert = lInsert & M46V999.gfForValor(1, flTabela.Value) & ", "
                            End If
                        Case adChar, adVarChar, adLongVarChar                        'Texto.
                            If M46V999.gfPreenchido(flTabela.Value) Then
                                lInsert = lInsert & "'" & M46V999.gfForAspas(flTabela.Value) & "', "
                            Else
                                lInsert = lInsert & "' ', "
                            End If
                        Case adBoolean                                  'Sim/N�o.
                            If IsNull(flTabela.Value) Then
                                lInsert = lInsert & "1, "
                            Else
                                lInsert = lInsert & "0, "
                            End If
                        Case Else
                            Call M46V999.gpFecha2(bdBasDados)
                            pMensagem = "FTE0008 - Tabela " & lIdeTabela & " com tipo do campo " & _
                                        flTabela.Name & " diferente dos previstos."
                            Exit Function
                        End Select
                    End Select
                Next flTabela
                If Mid(lInsert, Len(lInsert), 1) = "," Then
                    lInsert = Left(lInsert, Len(lInsert) - 1) & ")"
                Else
                    lInsert = Left(lInsert, Len(lInsert) - 2) & ")"
                End If
                If M46V999.gfExeSQL(bdBasDados, lInsert, pMensagem) <> 0 Then
                    Call M46V999.gpFecha2(bdBasDados)
                    GoTo ErrRollBack
                End If
                rsTabela.MoveNext
            Wend
        End If
    Next lCount
    lInsert = " Insert into Tab_Ped_Num_Copia values ("
    lInsert = lInsert & "'" & pNovNosNumero & "',"
    lInsert = lInsert & "'" & pNossoNumero & "',"
    lInsert = lInsert & Left(M46V999.gfForData(3), 8) & ")"

    Dummy = M46V999.gfExeSQL(bdBasDados, lInsert, pMensagem)

    Dummy = M46V999.gfExeSQL(bdBasDados, "Commit Transaction ", pMensagem)
    Set bdBasDados = Nothing
    Set flTabela = Nothing
    Set rsTabela = Nothing


    gfImportarTransmissao = True
    Exit Function
ErrRollBack:

    Call M46V999.gpGraLog(0, "gfImportarTransmissao ERRO Back - " & pMensagem)
    Dummy = M46V999.gfExeSQL(bdBasDados, "Rollback Transaction ", pMensagem)
    Set bdBasDados = Nothing
    Set flTabela = Nothing
    Set rsTabela = Nothing
    Exit Function

ControleErro:
    Call M46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call M46V999.gpGraLog(0, "gfImportarTransmissao ERRO - " & Err & " " & Err.Description)
    Set bdBasDados = Nothing
    Set flTabela = Nothing
    Set rsTabela = Nothing
End Function

Public Function gfSegNovo(ByRef pNossoNumero As String, pMensagem As String, pCod_Ramo As Integer, _
                          pCodSegmentacao As Integer, pCodPlano As Long, tagDynaTrace As String) As Boolean

' Optional ByRef pNossoNumero As String = ""
' Optional pMensagem As String = ""
' Optional pCod_Ramo As Integer = 0
' Optional pCodSegmentacao As Integer = 0
' Optional pCodPlano As Long = 0

'Fun��o: cria cota��o - seguro novo.

    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "gfSegNovo"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Dim bdP00MPED As ADODB.Connection    'Base de dados: P00MPED.
    Dim lCodProduto As Integer      'C�digo do produto.
    Dim lCodProtocolo As Long           'C�digo do protocolo.
    Dim lCodSequencia As Long           'C�digo da seq��ncia.
    Dim lDesTarifa As String      'Descri��o da tarifa.
'Dim pmensagem       As String           'Mensagem.
    Dim lNosNumero As String      'Nosso n�mero.
    Dim lNumPedido As String      'N�mero do pedido.
    Dim lVigDatIni As Double      'Data de vig�ncia: in�cio.
    Dim lVigDatTer As Double      'Data de vig�ncia: t�rmino.
    Dim lObjstcPed As stcA46V702B

    gfSegNovo = False
    On Error GoTo ControleErro
    M46V111.gCod_Dizer = 0
    '1. Abre base de dados.
    If M46V999.gfAbrBasDados("P00MPED", M46V999.gTipP00MPED, M46V999.gEndP00MPED, bdP00MPED, pMensagem) = False Then
        Call M46V999.gpGraLog(2, pMensagem)
        Exit Function
        'End
    End If

    '2. Prepara campos.
    'a) Nosso n�mero.
    If pNossoNumero = "" Then
        If M46V999.gfCalNosNumero(lNosNumero, pMensagem) = False Then
            Call M46V999.gpGraLog(2, pMensagem)
            Exit Function
            'End
        End If
        pNossoNumero = lNosNumero
    Else
        lNosNumero = pNossoNumero
    End If

    'b) N�mero do pedido.
    If M46V999.gfCalNumPedido(lNumPedido, pMensagem, bdP00MPED) = False Then
        Call M46V999.gpGraLog(2, pMensagem)
        Exit Function
    End If
    M46V999.gPrefixo = "Cota��o " & lNumPedido & ". "

    'c) Per�odo de vig�ncia.
    lVigDatIni = M46V999.gHoje
    If Right(lVigDatIni, 4) = "0229" Then
        lVigDatIni = M46V999.gHoje - 1
    End If
    lVigDatTer = lVigDatIni + 10000

    If Not IsNumeric(pCod_Ramo) Then pCod_Ramo = 0

    'Argumento passado na chamada do EXE se for igual a DEV � a vers�o de homologa��o
'    If InStr(UCase(Command), "DEV") = 0 And M46V999.gPrograma <> "D046O003" Then
'        If M46V999.gTarifa_Apolice < 1080 Then
'            lCodProduto = 1076
'        End If
'    End If

    'd) Produto, protocolo e seq��ncia.
    If M46V999.gfBusProduto(lVigDatIni, lCodProduto, lCodProtocolo, lCodSequencia, lDesTarifa, pMensagem, pCod_Ramo) = _
       False Then
        Call M46V999.gpGraLog(1, pMensagem)
        Exit Function
    End If

    '3. Grava registro de dados comuns.
    If gfGraNovPedido(bdP00MPED, lNosNumero, lNumPedido, lVigDatIni, lVigDatTer, _
                      lCodProduto, lCodProtocolo, lCodSequencia, pMensagem, pCod_Ramo, pCodPlano, pCodSegmentacao, lObjstcPed, tagDynaTrace) = False Then
        Call M46V999.gpGraLog(1, pMensagem)
        Exit Function
    End If

    '4. Grava registro de item.
    If M46V999.gfGraItem(bdP00MPED, lNosNumero, lNumPedido, "0001", pMensagem, 100, , , , , , pCod_Ramo, lCodProduto, pCodPlano, lObjstcPed) = False Then
        Call M46V999.gpGraLog(1, pMensagem)
        Exit Function
    End If

    '5. Grava registro de comunica��o.
    If M46V111.gfGraRegComunicacao(lNosNumero, pMensagem, bdP00MPED) = False Then
        Call M46V999.gpGraLog(2, pMensagem)
        Exit Function
    End If

    Set bdP00MPED = Nothing
    Set lObjstcPed = Nothing

    gfSegNovo = True
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
    
    Exit Function
ControleErro:
    Set bdP00MPED = Nothing
    Set lObjstcPed = Nothing
    Call M46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "gfSegNovo ERRO - " & Err & " " & Err.Description & " " & pMensagem, 2, Class_Name, methodName)
    Call M46V999.gpGraLog(0, "gfSegNovo ERRO - " & Err & " " & Err.Description & " " & pMensagem)

End Function
Public Function mfCarregaTriagem(ByRef pobjstcPed As stcA46V702B, ByRef pMensagem As String, tagDynaTrace As String) As Boolean
    
    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "mfCarregaTriagem"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    
    Dim lbdBasDados As ADODB.Connection    'Base de dados.
    Dim lrsTabela As clsYasRecordSet    'Registro da tabela.
    Dim lSql As String
    Dim lI As Integer
'    Dim lNum_pi As Long
    Dim lNum_pi As Double
    Dim lNosNumero As String
    Dim lPropTabela As String
    Dim ObjStcCorretor As stcA46V705B      'Objeto da classe de estrutura stcA46V705B (TAB_PED_CORR)
    Dim ObjCorretor As clsA46V705A   'Objeto da classe de estrutura stcA46V705B (TAB_PED_CORR)

    Set ObjCorretor = New clsA46V705A
    Set ObjCorretor.M46V999 = M46V999

    On Error GoTo mfCarregaTriagem_Err

    Call M46V999.gfObtSYASMULTINI(lPropTabela)

    mfCarregaTriagem = False

    lNum_pi = pobjstcPed.NUM_PI
    lNosNumero = pobjstcPed.nosso_numero

    If val(lNum_pi) = 0 Then
        mfCarregaTriagem = True
        Exit Function
    End If

    If M46V999.gfAbrBasDados("GED", M46V999.gTipGed, M46V999.gEndGed, lbdBasDados, pMensagem) = False Then Exit Function

    lSql = "Select * "
    lSql = lSql & " from tab_triagem WITH (NOLOCK)  where num_pi = " & lNum_pi

    If M46V999.gfObtRegistro(lbdBasDados, lSql, lrsTabela, pMensagem) = False Then
        Call M46V999.gpFecha2(lbdBasDados)
        Exit Function
    End If

    If lrsTabela.EOF Then
        mfCarregaTriagem = True
        GoTo Finalizar
    End If

    pobjstcPed.nosso_numero = lNosNumero
    pobjstcPed.Num_Apol_An = M46V999.Check(lrsTabela("num_apol_ant"))
    pobjstcPed.Dat_Ini_Vig = M46V999.Check(lrsTabela("dat_ini_vig"))
    pobjstcPed.Dat_Ter_Vig = M46V999.Check(lrsTabela("dat_fim_vig"))
    pobjstcPed.Cod_Unid_Prod = M46V999.Check(lrsTabela("cod_unidade"))
    pobjstcPed.Cod_Prod = M46V999.Check(lrsTabela("cod_produtor"))
    pobjstcPed.Cod_Ramo = M46V999.Check(lrsTabela("Cod_Ramo"))

    pobjstcPed.objstcCtrlEmis.Num_Bloq = M46V999.Check(lrsTabela("num_bloq"))
    If M46V999.Check(lrsTabela("num_bloq")) <> 0 Then
        pobjstcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec = e_SimNao_Sim
    End If

    If M46V999.Check(lrsTabela("cod_bco_car")) <> 0 And pobjstcPed.Cod_Bco_Cobr = 0 Then
        pobjstcPed.Cod_Bco_Cobr = M46V999.Check(lrsTabela("cod_bco_deb"))
    End If

    If M46V999.Check(lrsTabela("cod_bco_deb")) <> 0 And pobjstcPed.Cod_Bco_Cobr = 0 Then
        pobjstcPed.Cod_Bco_Cobr = M46V999.Check(lrsTabela("cod_bco_deb"))
        pobjstcPed.Cod_Agen_Bas_Cor = M46V999.Check(lrsTabela("cod_agenc"))
        pobjstcPed.Num_Cta_Bas_Cor = Left(Format(M46V999.Check(lrsTabela("cod_conta")), "00000000"), 7)
    End If

    pobjstcPed.objstcCtrlEmis.Tip_Cotac = M46V999.Check(lrsTabela("tip_cotac"))
    If pobjstcPed.objstcCtrlEmis.Tip_Cotac = "2" Then
        pobjstcPed.objstcCtrlEmis.Num_Cotac = M46V999.Check(lrsTabela("num_siscota"))
    End If
    pobjstcPed.objstcCtrlEmis.Dat_Ini_Vig = M46V999.Check(lrsTabela("dat_ini_vig"))
    pobjstcPed.objstcCtrlEmis.Dat_Ter_Vig = M46V999.Check(lrsTabela("dat_fim_vig"))
    pobjstcPed.objstcCtrlEmis.Cod_Corr = M46V999.Check(lrsTabela("cod_corr"))
    pobjstcPed.objstcCtrlEmis.Cod_Corr_2 = M46V999.Check(lrsTabela("cod_corr2"))
    pobjstcPed.objstcCtrlEmis.Cod_Corr_3 = M46V999.Check(lrsTabela("cod_corr3"))
    pobjstcPed.objstcCtrlEmis.COD_UNID = M46V999.Check(lrsTabela("cod_unidade"))
    pobjstcPed.objstcCtrlEmis.Cod_Prod = M46V999.Check(lrsTabela("cod_produtor"))
    pobjstcPed.objstcCtrlEmis.Cod_Ramo = M46V999.Check(lrsTabela("Cod_Ramo"))



    If pobjstcPed.ColecaoPedCorr.Count = 0 Then
        lI = 0
        If M46V999.gfPreenchido(lrsTabela("cod_corr")) Then
            Set ObjStcCorretor = New stcA46V705B
            Set ObjStcCorretor.M46V999 = M46V999

            ObjStcCorretor.Cod_Corr = lrsTabela("cod_corr")
            ObjStcCorretor.nosso_numero = pobjstcPed.nosso_numero
            ObjStcCorretor.Num_Ordem_Corr = lI + 1
            ObjStcCorretor.Cod_Ramo = pobjstcPed.Cod_Ramo
            ObjStcCorretor.Num_Ped = pobjstcPed.Num_Ped
            If Not ObjCorretor.mfIncluir(ObjStcCorretor, lPropTabela, "", tagDynaTrace) Then
                pMensagem = ObjCorretor.UltimoErro
            End If
            Set ObjStcCorretor = Nothing
            lI = lI + 1
        End If
        If M46V999.gfPreenchido(lrsTabela("cod_corr2")) Then
            Set ObjStcCorretor = New stcA46V705B
            Set ObjStcCorretor.M46V999 = M46V999

            ObjStcCorretor.Cod_Corr = lrsTabela("cod_corr2")
            ObjStcCorretor.nosso_numero = pobjstcPed.nosso_numero
            ObjStcCorretor.Num_Ordem_Corr = lI + 1
            ObjStcCorretor.Cod_Ramo = pobjstcPed.Cod_Ramo
            ObjStcCorretor.Num_Ped = pobjstcPed.Num_Ped
            If Not ObjCorretor.mfIncluir(ObjStcCorretor, lPropTabela, "", tagDynaTrace) Then
                pMensagem = ObjCorretor.UltimoErro
            End If
            lI = lI + 1
            Set ObjStcCorretor = Nothing
        End If
        If M46V999.gfPreenchido(lrsTabela("cod_corr3")) Then
            Set ObjStcCorretor = New stcA46V705B
            Set ObjStcCorretor.M46V999 = M46V999

            ObjStcCorretor.Cod_Corr = lrsTabela("cod_corr3")
            ObjStcCorretor.nosso_numero = pobjstcPed.nosso_numero
            ObjStcCorretor.Num_Ordem_Corr = lI + 1
            ObjStcCorretor.Cod_Ramo = pobjstcPed.Cod_Ramo
            ObjStcCorretor.Num_Ped = pobjstcPed.Num_Ped
            If Not ObjCorretor.mfIncluir(ObjStcCorretor, lPropTabela, "", tagDynaTrace) Then
                pMensagem = ObjCorretor.UltimoErro
            End If
            lI = lI + 1
            Set ObjStcCorretor = Nothing
        End If
    Else
        For Each ObjStcCorretor In pobjstcPed.ColecaoPedCorr
            If ObjStcCorretor.Cod_Corr <> lrsTabela("cod_corr") Or _
               ObjStcCorretor.Cod_Corr <> lrsTabela("cod_corr2") Or _
               ObjStcCorretor.Cod_Corr <> lrsTabela("cod_corr3") Then
                'Definir o que vai fazer quando existir corretor cadastrado

                'cod_corr,
                'cod_corr2,
                'cod_corr3
            End If
        Next
    End If


    mfCarregaTriagem = True
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)

Finalizar:
    Set lbdBasDados = Nothing
    Set lrsTabela = Nothing
    Set ObjStcCorretor = Nothing
    Set ObjCorretor = Nothing
    Set lbdBasDados = Nothing
    Set lrsTabela = Nothing
    Exit Function

mfCarregaTriagem_Err:
    pMensagem = "mfCarregaTriagem ERRO - " & Err & " " & Err.Description
    Call M46V999.gpGraLog(0, pMensagem)
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, pMensagem, 2, Class_Name, methodName)
    GoTo Finalizar
End Function

Public Sub gpAtualizaIni()
    Dim lPathPrinc As String
    Dim lPathArquivos As String
    Dim Aux As String
    Dim lMantem As Boolean

    If M46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then Exit Sub

    Aux = ""
    Aux = LerIni("MULTPROB_V0", "NOME", M46V999.gAppPath & "SyasMult.INI")
    lMantem = True
    If Trim(Aux) <> "" Then
        lPathArquivos = ""
        If InStrRev(Aux, "\") > 0 Then
            lPathArquivos = Left(Aux, InStrRev(Aux, "\"))
        End If
        If M46V999.gfPreenchido(lPathArquivos) = True And (M46V999.gTipUsuario = e_Tipo_Usu_Corretor Or M46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda) Then
            'Vai corrigir o path alterado indevidamente na vers�o de 28/09/2007
            lMantem = False
        Else
            Call M46V999.gpGraLog(0, "gpAtualizaIni: saiu sem reescrever")
            Exit Sub  'Se j� existir desconsidera
        End If
    End If

    lPathPrinc = LerIni("MULTCOT1", "CONFIG", M46V999.gAppPath & "SyasMult.INI")
    lPathArquivos = ""
    If InStrRev(lPathPrinc, "\") > 0 Then
        lPathArquivos = Left(lPathPrinc, InStrRev(lPathPrinc, "\"))
    End If

    Call M46V999.gpGraLog(0, "gpAtualizaIni: Caminho encontrado = " & lPathArquivos)
    If lMantem = False Then
        lPathArquivos = ""
        Call M46V999.gpGraLog(0, "gpAtualizaIni: corrigindo caminho = " & lPathArquivos)
    End If
    Call M46V999.gpGraLog(0, "gpAtualizaIni: Atualizando caminho CONFIG =" & lPathArquivos)

    Call GravarINI("MULTCOT1", "CONFIG", lPathArquivos & "MULTCOT1.CXY", M46V999.gAppPath & "SyasMult.INI")
    Call GravarINI("MULTCOT1", "IMAGEM", lPathArquivos & "MULTCOT1.WMF", M46V999.gAppPath & "SyasMult.INI")

    Call GravarINI("MULTCOT2", "CONFIG", lPathArquivos & "MULTCOT2.CXY", M46V999.gAppPath & "SyasMult.INI")
    Call GravarINI("MULTCOT2", "IMAGEM", lPathArquivos & "MULTCOT2.WMF", M46V999.gAppPath & "SyasMult.INI")

    Call GravarINI("MULTCOT3", "CONFIG", lPathArquivos & "MULTCOT3.CXY", M46V999.gAppPath & "SyasMult.INI")
    Call GravarINI("MULTCOT3", "IMAGEM", lPathArquivos & "MULTCOT3.WMF", M46V999.gAppPath & "SyasMult.INI")

    Call GravarINI("MULTCOTA", "CONFIG", lPathArquivos & "MULTCOTA.CXY", M46V999.gAppPath & "SyasMult.INI")
    Call GravarINI("MULTCOTA", "IMAGEM", lPathArquivos & "MULTCOTA.WMF", M46V999.gAppPath & "SyasMult.INI")

    Call GravarINI("MULTCOTC", "CONFIG", lPathArquivos & "MULTCOTC.CXY", M46V999.gAppPath & "SyasMult.INI")
    Call GravarINI("MULTCOTC", "IMAGEM", lPathArquivos & "MULTCOTC.WMF", M46V999.gAppPath & "SyasMult.INI")

    Call GravarINI("MULTPRO1", "CONFIG", lPathArquivos & "MULTPRO1.CXY", M46V999.gAppPath & "SyasMult.INI")
    Call GravarINI("MULTPRO1", "IMAGEM", lPathArquivos & "MULTPRO1.WMF", M46V999.gAppPath & "SyasMult.INI")

    Call GravarINI("MULTPRO2", "CONFIG", lPathArquivos & "MULTPRO2.CXY", M46V999.gAppPath & "SyasMult.INI")
    Call GravarINI("MULTPRO2", "IMAGEM", lPathArquivos & "MULTPRO2.WMF", M46V999.gAppPath & "SyasMult.INI")

    Call GravarINI("MULTPRO3", "CONFIG", lPathArquivos & "MULTPRO3.CXY", M46V999.gAppPath & "SyasMult.INI")
    Call GravarINI("MULTPRO3", "IMAGEM", lPathArquivos & "MULTPRO3.WMF", M46V999.gAppPath & "SyasMult.INI")

    Call GravarINI("MULTPRO4", "CONFIG", lPathArquivos & "MULTPRO4.CXY", M46V999.gAppPath & "SyasMult.INI")
    Call GravarINI("MULTPRO4", "IMAGEM", lPathArquivos & "MULTPRO4.WMF", M46V999.gAppPath & "SyasMult.INI")

    Call GravarINI("MULTPROB", "CONFIG", lPathArquivos & "MULTPROB.CXY", M46V999.gAppPath & "SyasMult.INI")
    Call GravarINI("MULTPROB", "IMAGEM", lPathArquivos & "MULTPROB.WMF", M46V999.gAppPath & "SyasMult.INI")

    Call GravarINI("MULTPROC", "CONFIG", lPathArquivos & "MULTPROC.CXY", M46V999.gAppPath & "SyasMult.INI")
    Call GravarINI("MULTPROC", "IMAGEM", lPathArquivos & "MULTPROC.WMF", M46V999.gAppPath & "SyasMult.INI")

    Call GravarINI("MULTPROD", "CONFIG", lPathArquivos & "MULTPROD.CXY", M46V999.gAppPath & "SyasMult.INI")
    Call GravarINI("MULTPROD", "IMAGEM", lPathArquivos & "MULTPROD.WMF", M46V999.gAppPath & "SyasMult.INI")


    Call GravarINI("MULTCOT1_V0", "NOME", lPathArquivos & "MULTCOT1_V0", M46V999.gAppPath & "SyasMult.INI")
    Call GravarINI("MULTCOT2_V0", "NOME", lPathArquivos & "MULTCOT2_V0", M46V999.gAppPath & "SyasMult.INI")
    Call GravarINI("MULTCOT3_V0", "NOME", lPathArquivos & "MULTCOT3_V0", M46V999.gAppPath & "SyasMult.INI")
    Call GravarINI("MULTPRO1_V0", "NOME", lPathArquivos & "MULTPRO1_V0", M46V999.gAppPath & "SyasMult.INI")
    Call GravarINI("MULTPRO2_V0", "NOME", lPathArquivos & "MULTPRO2_V0", M46V999.gAppPath & "SyasMult.INI")
    Call GravarINI("MULTPRO3_V0", "NOME", lPathArquivos & "MULTPRO3_V0", M46V999.gAppPath & "SyasMult.INI")
    Call GravarINI("MULTPRO4_V0", "NOME", lPathArquivos & "MULTPRO4_V0", M46V999.gAppPath & "SyasMult.INI")
    Call GravarINI("MULTPROB_V0", "NOME", lPathArquivos & "MULTPROB_V0", M46V999.gAppPath & "SyasMult.INI")

End Sub

Private Function mfGraRenItem_Roubo(prstItem As clsYasRecordSet, pNosNumero As String, _
                                    pBdP00APOL As ADODB.Connection, pMensagem As String, _
                                    pNumPed As Long, pCod_Dizer As Integer, tagDynaTrace As String) As Boolean
'Fun��o: Grava os registros de comum com os dados da ap�lice (Renova��o Yasuda).
'Par�metros de entrada :prstItem.........Registro de item(ap�lice)
'                       pNosNumero.......Nosso n�mero.
'                       pBdP00APOL.......Arquivo de conex�o
'Par�metro de sa�da:    pMensagem........Mensagem de erro.

    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "mfGraRenItem_Roubo"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Dim lSql As String        'Consulta SQL.
    Dim rsAPOL As clsYasRecordSet    'Registro.

'[FRH]->
'Conferir coberturas na tabela POOMULTI:TAB_COB_RELACAO
    Dim lSql_buscaCob As String
    Dim rs_buscaCob As clsYasRecordSet
    Dim bdP00MULTI As ADODB.Connection
    '<-[FRH]

    Dim lobjPedLoc As clsA46V708A         'Objeto da classe clsA46V708 (TAB_PED_LOC).
    Dim lobjstcPedLoc As stcA46V708B            'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC).
    Dim lObjClau As clsA46V703A       'Objeto da classe clsA46V703A (TAB_PED_CLAU).
    Dim lobjstcClau As stcA46V703B          'Objeto da classe de estrutura stcA46V703B (TAB_PED_CLAU).
    Dim lobjCobert As clsA46V704A         'Objeto da classe clsA46V704A (TAB_PED_COBERT).
    Dim lobjstcCobert As stcA46V704B            'Objeto da classe de estrutura stcA46V704B (TAB_PED_COBERT).
    Dim lobjInsp As clsA46V717A       'Objeto da classe clsA46V717A (TAB_PED_INSP).
    Dim lObjstcInsp As stcA46V717B          'Objeto da classe de estrutura stcA46V717B (TAB_PED_INSP).
    Dim lNumero As String
    Dim lComplemento As String
    Dim lPos As Long
    Dim lNomeTab As String
    Dim lValorRisco As Double
    Dim lObjPedFator As clsA46V720A
    Dim lobjstcPedFator As stcA46V720B
    Dim lPropTabela As String

    On Error GoTo ControleErro

    Call M46V999.gfObtSYASMULTINI(lPropTabela)

    If M46V999.gBaseLocal = True Then
        lNomeTab = "TAB_APOL"
    Else
        lNomeTab = "GED.DBO.TAB_PED"
    End If

    Set lobjPedLoc = New clsA46V708A
    Set lobjPedLoc.M46V999 = M46V999

    Set lobjstcPedLoc = New stcA46V708B
    Set lobjstcPedLoc.M46V999 = M46V999

    With lobjstcPedLoc
        .nosso_numero = pNosNumero
        .Num_Ped = pNumPed
        If M46V111.gNum_Pi = 0 Then
            .NUM_PI = Right(Year(Date), 2) & Format(pNumPed, "00000000")
        Else
            .NUM_PI = M46V111.gNum_Pi
        End If
        .Cod_Ramo = M46V999.gfTratarNulo(prstItem("Cod_Ramo"), tpoNumerico)
        .Num_Item = 1
        .Per_Desc_Apar = 0
        .Per_Desc_Progres = 0
        .Per_IS_Vr = 0
        .Val_Coef_IS_Vr = 0
        .Val_Coef_IS_Vr_LC = 0
        .Val_Vr = M46V999.gfTratarNulo(prstItem("Val_IS"), TpoValor)
        lValorRisco = .Val_Vr
        .Tab_Clasf = 0
        .Cod_Clasf = ""
        .Cod_Compl_Clasf = ""
        .Dsc_Clasf = ""
        .Cod_Constr = 0
        .Cod_User_Liber = M46V999.gNomUsuario6
        .Dat_Liber = Left(M46V999.gfForData(3), 8)
        .Hor_Liber = Mid(M46V999.gfForData(3), 9, 4)
        .Cod_Sit_item = e_SitItem_NaoCalc
        .Per_Bonus = 0
        .Num_Apol_Anterior = prstItem("Num_APol")
        .Num_Item_Anterior = 1
        If lobjstcPedLoc.mBonus_Anterior = 0 Then
            .Per_Bonus = lobjPedLoc.mfRetornaBonusApolice(lobjstcPedLoc, tagDynaTrace)
        Else
            .Per_Bonus = lobjstcPedLoc.mBonus_Anterior
        End If
        .Ind_Sinistro = lobjstcPedLoc.Apol_Ant_IND_SINISTRO
        If .Ind_Sinistro = 0 Then .Ind_Sinistro = e_SimNao_Nao

        If .Ind_Sinistro = e_SimNao_Sim Then
            .Per_Bonus = .Per_Bonus - 5
        Else
            .Per_Bonus = .Per_Bonus + 5
        End If
        If .Per_Bonus < 0 Then .Per_Bonus = 0
        If .Per_Bonus > 25 Then .Per_Bonus = 25
        M46V111.gDesc_Bonus = .Per_Bonus
        .Per_Tot_Protec = .gPer_Tot_Protec
        .Cod_Produto = M46V111.gCod_plano
        .Cod_Regiao = M46V999.gfTratarNulo(prstItem("Cod_Reg"), tpoNumerico)
        If pCod_Dizer = 0 Then
            .Tip_Emissao = e_TipEmissao_RenYasuda
        Else
            .Tip_Emissao = e_TipEmissao_Endosso
        End If
        .Cod_Iden_Tarif = M46V999.gCodProduto

    End With

    If Not lobjPedLoc.mfIncluir(lobjstcPedLoc, lPropTabela, "", tagDynaTrace) Then
        pMensagem = lobjPedLoc.UltimoErro
        GoTo SAIDA
    End If

    'Clausulas
    lSql = " Select * from " & lNomeTab & "_CLAU WITH (NOLOCK) "
    If M46V999.gBaseLocal = False And pBdP00APOL Like "*Provider=SQLOLEDB.*" Then
        lSql = lSql & " inner join GED.DBO.tab_ctrl_emis  as ctrl WITH (NOLOCK) on " & lNomeTab & "_CLAU.num_apol = ctrl.num_apol"
        lSql = lSql & " and ctrl.num_endo = " & lNomeTab & "_CLAU.num_endo"
        lSql = lSql & " and ctrl.num_PI = " & lNomeTab & "_CLAU.num_PI"
    End If
    lSql = lSql & " Where " & lNomeTab & "_CLAU.Num_Apol = " & prstItem("Num_Apol")
    lSql = lSql & " and " & lNomeTab & "_CLAU.num_item = " & 1

    If Not M46V999.gfObtRegistro(pBdP00APOL, lSql, rsAPOL, pMensagem) Then
        GoTo ControleErro
    End If

    Set lObjClau = New clsA46V703A
    Set lObjClau.ConexaoBD = lobjPedLoc.ConexaoBD
    Set lObjClau.M46V999 = M46V999

    Do While Not rsAPOL.EOF
        Set lobjstcClau = New stcA46V703B
        Set lobjstcClau.M46V999 = M46V999

        With lobjstcClau
            .Cod_Ramo = M46V999.gfTratarNulo(rsAPOL("Cod_Ramo"), tpoNumerico)
            .nosso_numero = pNosNumero
            .Num_Ped = pNumPed
            If M46V111.gNum_Pi = 0 Then
                .NUM_PI = Right(Year(Date), 2) & Format(pNumPed, "00000000")
            Else
                .NUM_PI = M46V111.gNum_Pi
            End If
            .Cod_Clau = M46V999.gfTratarNulo(rsAPOL("Cod_Clau"), tpoNumerico)
            .Num_Item = lobjstcPedLoc.Num_Item
        End With
        If Not lObjClau.mfIncluir(lobjstcClau, lPropTabela, "", tagDynaTrace) Then
            pMensagem = lObjClau.UltimoErro
            GoTo SAIDA
        End If
        rsAPOL.MoveNext
    Loop

    'COberturas
    lSql = " Select " & lNomeTab & "_COBERT.* from " & lNomeTab & "_COBERT WITH (NOLOCK) "
    If M46V999.gBaseLocal = False And pBdP00APOL Like "*Provider=SQLOLEDB.*" Then
        lSql = lSql & " inner join GED.DBO.tab_ctrl_emis  as ctrl WITH (NOLOCK) on GED.DBO.TAB_PED_COBERT.num_apol = ctrl.num_apol"
        lSql = lSql & " and ctrl.num_endo = GED.DBO.TAB_PED_COBERT.num_endo"
        lSql = lSql & " and ctrl.num_PI = GED.DBO.TAB_PED_COBERT.num_PI"
    End If
    lSql = lSql & " Where " & lNomeTab & "_COBERT.Num_Apol = " & prstItem("Num_Apol")
    lSql = lSql & " and " & lNomeTab & "_COBERT.num_item = " & 1

    If Not M46V999.gfObtRegistro(pBdP00APOL, lSql, rsAPOL, pMensagem) Then
        GoTo ControleErro
    End If

    Set lobjCobert = New clsA46V704A
    Set lobjCobert.M46V999 = M46V999

    'Set lObjCobert.ConexaoBD = lObjPedLoc.ConexaoBD
    '[FRH]->
    If Not M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, bdP00MULTI, pMensagem) Then
        GoTo ControleErro
    End If
    '<-[FRH]
    Do While Not rsAPOL.EOF
        Set lobjstcCobert = New stcA46V704B
        Set lobjstcCobert.M46V999 = M46V999

        With lobjstcCobert
            .Cod_Ramo = M46V999.gfTratarNulo(rsAPOL("Cod_Ramo"), tpoNumerico)
            .nosso_numero = pNosNumero
            .Num_Ped = pNumPed
            If M46V111.gNum_Pi = 0 Then
                .NUM_PI = Right(Year(Date), 2) & Format(pNumPed, "00000000")
            Else
                .NUM_PI = M46V111.gNum_Pi
            End If
            .Num_Item = lobjstcPedLoc.Num_Item

            '[FRH]->
            'Rotina para conferir altera��es de c�digos de cobertura,
            'na *POOMULTI.TAB_COB_RELACAO* existe um de-para de cobs.anteriores e atuais.

            '[Valor original Cod_Cobert da TAB_APOL]
            .Cod_Cobert = M46V999.gfTratarNulo(rsAPOL("cod_cobert"), tpoNumerico)

            lSql_buscaCob = "SELECT COD_COB_ATUAL FROM TAB_COB_RELACAO WITH (NOLOCK) WHERE COD_COB_ANTERIOR=" & M46V999.gfTratarNulo(rsAPOL("Cod_Cobert"), tpoNumerico)
            If M46V999.gfObtRegistro(bdP00MULTI, lSql_buscaCob, rs_buscaCob, pMensagem) Then
                If Not (rs_buscaCob.BOF And rs_buscaCob.EOF) Then
                    .Cod_Cobert = M46V999.gfTratarNulo(rs_buscaCob("cod_cob_atual"), tpoNumerico)
                End If
            End If
            '<-[FRH]

            '    .Tab_Cobert = M46V999.gfTratarNulo(rsApol("Tab_Cobert"), tpoNumerico)
            If M46V111.gCod_plano > 1010 Then
                .Tab_Cobert = M46V111.gCod_plano
            Else
                .Tab_Cobert = 1
            End If


            .Per_Desc_Exp = M46V111.gDesc_Bonus

            .Val_IS = M46V999.gfTratarNulo(rsAPOL("Val_IS"), tpoNumerico)
            .Val_Risco_Cob = .Val_IS

            If pCod_Dizer <> 0 Then
                If M46V111.gCod_plano > 1010 And .Cod_Ramo = e_CodRamo_Residencial Then
                    If InStr("3,6", pCod_Dizer) <> 0 Then
                        .Alterar_Cobert = "FIXO"
                    ElseIf InStr("2,4,10,11,18", pCod_Dizer) <> 0 Then
                        .Alterar_Cobert = "FIXA IS"
                    End If
                ElseIf .Cod_Ramo = e_CodRamo_Residencial Then
                    If InStr("3,6", pCod_Dizer) <> 0 Then
                        .Alterar_Cobert = "FIXO"
                    End If
                Else
                    If InStr("2,3,4,5,6,7,10,11,18", pCod_Dizer) <> 0 Then
                        .Alterar_Cobert = "FIXO"
                    End If
                End If
            End If

            '11/01/2010 - Ficha: 1000003 - franquia e valor de franquia - inicio
            'Call mpBuscaLimFranquia(lObjstcPedLoc, lobjstcCobert)

            Dim lValFraqMinDefault As Double
            Dim lPerFranqDefault As Double
            Dim lFatorFranqDefault As Double
            Dim lValFraqMaxDefault As Double
            Dim TextoFranquia As String
            Dim DescrFranquia As String


            Call M46V999.mpNovoBuscaLimFranquia(lobjstcPedLoc.Cod_Ramo, lobjstcPedLoc.ObjPai.Cod_Estip, lobjstcPedLoc.ObjPai.Cod_Produto, _
                                                lobjstcCobert.Cod_Cobert, lobjstcPedLoc.Cod_Clasf, lobjstcPedLoc.Cod_Compl_Clasf, _
                                                lobjstcPedLoc.Cod_Constr, lobjstcPedLoc.Num_Cep_Risco, lobjstcPedLoc.ObjPai.Cod_Sequencia, lobjstcPedLoc.ObjPai.objstcCtrlEmis.Cod_Corr, _
                                                lValFraqMinDefault, lPerFranqDefault, lFatorFranqDefault, lValFraqMaxDefault, TextoFranquia, DescrFranquia, lobjstcCobert.Num_Meses_PI)

            lobjstcCobert.Val_Franq_Min = lValFraqMinDefault
            lobjstcCobert.Val_Franq_Max = lValFraqMaxDefault
            lobjstcCobert.Per_Franq = lPerFranqDefault
            lobjstcCobert.Val_Fator_Franq = lFatorFranqDefault
            lobjstcCobert.TextoFranquia = TextoFranquia

            '11/01/2010 - Ficha: 1000003 - franquia e valor de franquia - fim

        End With

        If Not lobjCobert.mfIncluir(lobjstcCobert, lPropTabela, "", False, False, tagDynaTrace) Then
            If UCase(lobjCobert.UltimoErro) Like "*DUPLICATE*" And Not rs_buscaCob.EOF And Not rs_buscaCob.BOF Then
                GoTo Proximo
            Else
                pMensagem = lobjCobert.UltimoErro
                GoTo SAIDA
            End If
        End If
        If lobjstcCobert.Cod_Cobert = 1 And lValorRisco = 0 Then
            lobjstcPedLoc.Val_Vr = lobjstcCobert.Val_IS
            lobjstcPedLoc.Per_IS_Vr = 100
            If Not lobjPedLoc.mfAlterarValorRisco(lobjstcPedLoc, lPropTabela, "") Then
                pMensagem = lobjPedLoc.UltimoErro
                GoTo SAIDA
            End If
        End If
Proximo:
        rsAPOL.MoveNext
    Loop

    '[FRH]->
    Call M46V999.gpFecha2(bdP00MULTI)
    '<-[FRH]

    Set lobjInsp = New clsA46V717A
    Set lobjInsp.M46V999 = M46V999

    Set lObjstcInsp = New stcA46V717B
    lObjstcInsp.nosso_numero = pNosNumero
    lObjstcInsp.Num_Item = lobjstcPedLoc.Num_Item
    lObjstcInsp.Num_Logr_Risco = val(lNumero)
    'corre��o
    lObjstcInsp.Cod_Ramo = lobjstcPedLoc.Cod_Ramo
    lObjstcInsp.Num_Ped = lobjstcPedLoc.Num_Ped
    '
    If Not lobjInsp.mfIncluir(lObjstcInsp, lPropTabela, "", tagDynaTrace) Then
        pMensagem = lobjInsp.UltimoErro
        GoTo SAIDA
    End If


    'Ficha 1000409 - Inclus�o de Fatores de Prote��o em Renova��es Yasuda
    If M46V999.gBaseLocal = True Then
        lNomeTab = "TAB_APOL"
    Else
        lNomeTab = "RamosDiversos.dbo.Tab_Ped"
    End If
    lSql = " Select * from " & lNomeTab & "_FAT_RISCO WITH (NOLOCK) "
    If M46V999.gBaseLocal = False And pBdP00APOL Like "*Provider=SQLOLEDB.*" Then
        lSql = lSql & " inner join GED.DBO.tab_ctrl_emis  as ctrl WITH (NOLOCK) on " & lNomeTab & "_FAT_RISCO.num_apol = ctrl.num_apol"
        lSql = lSql & " and ctrl.num_endo = " & lNomeTab & "_FAT_RISCO.num_endo"
        lSql = lSql & " and ctrl.num_PI = " & lNomeTab & "_FAT_RISCO.num_PI"
    End If
    lSql = lSql & " Where " & lNomeTab & "_FAT_RISCO.Num_Apol = " & prstItem("Num_Apol")
    lSql = lSql & " and " & lNomeTab & "_FAT_RISCO.num_item = " & 1
    '    lSql = lSql & " and " & lNomeTab & "_FAT_RISCO.num_Endo = " & prstItem("NUm_Endo2")

    On Error Resume Next
    If Not M46V999.gfObtRegistro(pBdP00APOL, lSql, rsAPOL, pMensagem) Then
        GoTo ControleErro
    End If

    Set lObjPedFator = New clsA46V720A
    Set lObjPedFator.M46V999 = M46V999

    Set lObjPedFator.ConexaoBD = lobjPedLoc.ConexaoBD
    Do While Not rsAPOL.EOF
        Set lobjstcPedFator = New stcA46V720B
        With lobjstcPedFator
            .Cod_Ramo = M46V999.gfTratarNulo(rsAPOL("Cod_Ramo"), tpoNumerico)
            .nosso_numero = pNosNumero
            .Num_Ped = pNumPed
            If M46V111.gNum_Pi = 0 Then
                .NUM_PI = Right(Year(Date), 2) & Format(pNumPed, "00000000")
            Else
                .NUM_PI = M46V111.gNum_Pi
            End If
            .Num_Item = lobjstcPedLoc.Num_Item
            .Cod_Fator = M46V999.gfTratarNulo(rsAPOL("Cod_Fator"), tpoNumerico)
            .Dsc_Compl_Fator = M46V999.gfTratarNulo(rsAPOL("DSC_COMPL_FATOR"), tpoNumerico)
        End With
        If Not lObjPedFator.mfIncluir(lobjstcPedFator, lPropTabela, "", tagDynaTrace) Then
            pMensagem = lObjPedFator.UltimoErro
            GoTo SAIDA
        End If
        rsAPOL.MoveNext
    Loop
    On Error GoTo ControleErro

    '************ Finalizado com Sucesso
    mfGraRenItem_Roubo = True
    '************
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)

SAIDA:
    GoSub FechaObj
    Exit Function

ControleErro:
    Call M46V999.gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call M46V999.gpGraLog(0, "gfGraRenItem ERRO - " & Err & " " & Err.Description & " " & pMensagem)
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "gfGraRenItem ERRO - " & Err & " " & Err.Description & " " & pMensagem, 2, Class_Name, methodName)
    GoTo SAIDA

FechaObj:
    Call M46V999.gpFecha2(bdP00MULTI)
    Set rsAPOL = Nothing
    Set rs_buscaCob = Nothing
    Set lobjPedLoc = Nothing
    Set lobjstcPedLoc = Nothing
    Set lObjClau = Nothing
    Set lobjstcClau = Nothing
    Set lobjCobert = Nothing
    Set lobjstcCobert = Nothing
    Set lobjInsp = Nothing
    Set lObjstcInsp = Nothing
    Set lObjPedFator = Nothing
    Set lobjstcPedFator = Nothing
    Return
End Function

Public Function gfRetornaCodProdutor(ByVal pCodCorr As Long, ByVal pCodUnidProd As Integer) As Integer
    Dim lCN As ADODB.Connection
    Dim lrs As clsYasRecordSet
    Dim lCodRetorno As Integer
    Dim lMsgErr As String
    Dim lSelect As String

    On Error GoTo TrataErro: gfRetornaCodProdutor = 0

    If Not M46V999.gPrograma = "D046O003" Then GoTo SAIDA

    If M46V999.gfAbrBasDados("P00MPED", M46V999.gTipP00MPED, M46V999.gEndP00MPED, lCN, _
                             lMsgErr) = False Then GoTo TrataErro

    lSelect = "EXEC [RamosDiversos].[dbo].[proc_RetornaProdutor] " & pCodCorr & ", " & pCodUnidProd

    If Not M46V999.gfObtRegistro(lCN, lSelect, lrs, lMsgErr) Then
        GoTo TrataErro
    ElseIf lrs.EOF = False Then
        lCodRetorno = M46V999.gfTratarNulo(lrs("Cod_Prod"), tpoNumerico)
    End If

    gfRetornaCodProdutor = lCodRetorno

SAIDA:
    GoSub FechaObj
    Exit Function
TrataErro:
    If Not M46V999.gfPreenchido(lMsgErr) Then
        lMsgErr = Err & " - " & ERROR
    End If
    lMsgErr = "gfRetornaCodProdutor Erro:" & lMsgErr
    Call M46V999.gpGraLog(0, lMsgErr)
    GoTo SAIDA
FechaObj:
    Call M46V999.gpFecha2(lCN)
    Set lrs = Nothing
    Set lCN = Nothing
    Return
End Function

Public Property Get M46V999() As clsM46V999
    Set M46V999 = M46V111.M46V999
End Property

Public Property Set M46V111(pClsM46V111 As clsM46V111)
    Set mClsM46V111 = pClsM46V111
End Property

Public Property Get M46V111() As clsM46V111
    If mClsM46V111 Is Nothing Then
        Err.Raise 91, Me.Class_Name, "Objeto M46V111 n�o atribu�do para a classe (" & Me.Class_Name & ")"
        '***Esta Classe DEVE utilizar uma instancia da classe clsM46V111 j�
        '   iniciada no escopo onde esta classe foi instanciada.
        'Atribua a o objeto (clsM46V111 ja instanciado) no momento que eu inicializou esta classe (atraves do Property SET)!!!
    End If

Retorno:
    Set M46V111 = mClsM46V111
End Property

Public Property Get M46V103() As clsM46V103
    Set M46V103 = M46V111.M46V103
End Property

Private Sub Class_Initialize()
    Set mClsM46V111 = Nothing
End Sub

Private Sub Class_Terminate()
    Set mClsM46V111 = Nothing
End Sub

Private Function mfRenConstrucaoDePara(ByVal Cod_Constr As Integer, ByVal Cod_Ramo As Integer, ByVal Cod_Produto As Integer) As Integer
    
    If Cod_Produto <= 1080 And Cod_Ramo = e_CodRamo_Empresarial And Cod_Constr = 2 Then
       Cod_Constr = 3
    
    Else
        If Cod_Produto <= 1080 And Cod_Ramo = e_CodRamo_Residencial And Cod_Constr = 2 Then
           Cod_Constr = 3
        End If
    End If
    
    mfRenConstrucaoDePara = Cod_Constr
End Function
