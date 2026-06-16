Attribute VB_Name = "basM46V001"
Option Explicit
Public gerro As Integer
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nsize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal filename$)

Function GravarINI(Sessao As String, NomeChave As String, Valor As String, NomedoArquivo As String)
Dim rc As Long
   rc = WritePrivateProfileString(Sessao, NomeChave, Valor, NomedoArquivo)
End Function

Function LerIni(Secao As String, NomeChave As String, NomedoArquivo As String) As String
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

Function mfCriaIndex(pMensagem As String) As Boolean
Dim lSql As String
Dim lCreate As String

       ' lSql = " Drop Index Tab_Ped_Cobert_Compl_IND03 on Tab_Ped_Cobert_Compl "
       ' If gfExeSQL(bdP00MPED, lSql, pMensagem) <> 0 Then
       '     Call gpGraLog(0, pMensagem)
       ' End If
       '
       ' lCreate = ""
       ' lCreate = lCreate & "CREATE  UNIQUE  INDEX   Tab_Ped_Cobert_Compl_IND33"
       ' lCreate = lCreate & "       On Tab_Ped_Cobert_Compl  (NOSSO_NUMERO, NUM_ITEM, COD_COBERT);"
       ' If gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
       '     Call gpGraLog(1, pMensagem)
       ' End If
       '
       ' lSql = " Drop Index Tab_Ped_Fat_Risco_IND03 on Tab_Ped_Fat_Risco "
       ' If gfExeSQL(bdP00MPED, lSql, pMensagem) <> 0 Then
       '     Call gpGraLog(0, pMensagem)
       ' End If
       '
       ' lCreate = ""
       ' lCreate = lCreate & "CREATE  UNIQUE  INDEX Tab_Ped_Fat_Risco_IND03"
       ' lCreate = lCreate & "       On Tab_Ped_Fat_Risco (NOSSO_NUMERO, NUM_ITEM, COD_FATOR);"
       ' If gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
       '     Call gpGraLog(1, pMensagem)
       ' End If
        '
       ' lSql = " Drop Index Tab_Ped_Espec2_IND03 on Tab_Ped_Espec2 "
       ' If gfExeSQL(bdP00MPED, lSql, pMensagem) <> 0 Then
       '     Call gpGraLog(0, pMensagem)
       ' End If
       '
       '
       '  lCreate = ""
       '  lCreate = lCreate & "CREATE  UNIQUE  INDEX Tab_Ped_Espec2_IND03"
       '  lCreate = lCreate & "       On Tab_Ped_Espec2 (NOSSO_NUMERO);"
       '  If gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
       '      Call gpGraLog(1, pMensagem)
       '  End If
       '
       '
       ' lSql = " Drop Index Tab_Ped_Espec_IND03 on Tab_Ped_Espec "
       ' If gfExeSQL(bdP00MPED, lSql, pMensagem) <> 0 Then
       '     Call gpGraLog(0, pMensagem)
       ' End If
       '
       ' lCreate = ""
       ' lCreate = lCreate & "CREATE  UNIQUE  INDEX Tab_Ped_Espec_IND03"
       ' lCreate = lCreate & "       On Tab_Ped_Espec (NOSSO_NUMERO, NUM_PAG);"
        'If gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
       '     Call gpGraLog(1, pMensagem)
        'End If'
'


End Function

Function mfCriaTabelas(pMensagem As String) As Boolean
On Error GoTo CriaTabelas_Err

Dim lCreate              As String
Dim bdP00MPED            As ADODB.Connection 'Base de dados: P00MPED.
Dim lrsTabela            As ADODB.Recordset  'Registro da tabela.
Dim lSql                 As String
Dim bdP00MULTI           As ADODB.Connection 'Base de dados: P00MULTI.


    If gTipUsuario = e_Tipo_Usu_Yasuda Then
        mfCriaTabelas = True
        Exit Function
    End If
    Screen.MousePointer = vbHourglass
    
    mfCriaTabelas = False
    '1. Abre base de dados.
    If gfAbrBasDados("P00MPED", gTipP00MPED, gEndP00MPED, bdP00MPED, pMensagem) = False Then
        Call gpGraLog(2, pMensagem)
        Exit Function
        'End
    End If
    
    lSql = "Select top 1 * from Tab_Ped_Cobert_Compl "
    pMensagem = ""
    If Not gfObtRegistro(bdP00MPED, lSql, lrsTabela, pMensagem) Then
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
        lCreate = lCreate & "     NUM_PI   double  NULL )  ;"
         
        If gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call gpGraLog(1, pMensagem)
        End If
         
        lSql = " Drop Index Tab_Ped_Cobert_Compl_IND03 on Tab_Ped_Cobert_Compl "
        If gfExeSQL(bdP00MPED, lSql, pMensagem) <> 0 Then
            Call gpGraLog(0, pMensagem)
        End If
         
         
        lCreate = ""
        lCreate = lCreate & "CREATE  UNIQUE  INDEX   Tab_Ped_Cobert_Compl_IND03"
        lCreate = lCreate & "       On Tab_Ped_Cobert_Compl  (NUM_PI, NUM_ITEM, COD_COBERT);"
        If gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call gpGraLog(1, pMensagem)
        End If
        
        lCreate = ""
        lCreate = lCreate & "CREATE    INDEX     Tab_Ped_Cobert_Compl_IND01"
        lCreate = lCreate & "       On Tab_Ped_Cobert_Compl (COD_ORG,COD_RAMO,NUM_PED) ;"
        If gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call gpGraLog(1, pMensagem)
        End If
        
        lCreate = ""
        lCreate = lCreate & "CREATE    INDEX   Tab_Ped_Cobert_Compl_IND02"
        lCreate = lCreate & "       On Tab_Ped_Cobert_Compl  (NUM_APOL,NUM_ENDO,NUM_ITEM);"
        If gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call gpGraLog(1, pMensagem)
        End If
    End If
    
    lSql = "Select top 1 * from Tab_Ped_Fat_Risco "
    pMensagem = ""
    If Not gfObtRegistro(bdP00MPED, lSql, lrsTabela, pMensagem) Then
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
        lCreate = lCreate & "     NUM_PI   long   NULL )  ;"
        If gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call gpGraLog(1, pMensagem)
        End If
         
        
        lCreate = ""
        lCreate = lCreate & "CREATE  UNIQUE  INDEX Tab_Ped_Fat_Risco_IND03"
        lCreate = lCreate & "       On Tab_Ped_Fat_Risco (NUM_PI, NUM_ITEM, COD_FATOR);"
        If gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call gpGraLog(1, pMensagem)
        End If
        
        lCreate = ""
        lCreate = lCreate & "CREATE    INDEX Tab_Ped_Fat_Risco_IND01"
        lCreate = lCreate & "       On Tab_Ped_Fat_Risco (COD_ORG,COD_RAMO,NUM_PED, COD_FATOR) ;"
        If gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call gpGraLog(1, pMensagem)
        End If
        
        lCreate = ""
        lCreate = lCreate & "CREATE    INDEX Tab_Ped_Fat_Risco_IND02"
        lCreate = lCreate & "       On Tab_Ped_Fat_Risco (NUM_APOL,NUM_ENDO,NUM_ITEM, COD_FATOR);"
        If gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call gpGraLog(1, pMensagem)
        End If
    End If
    
    lSql = "Select top 1 * from Tab_Ped_Espec2 "
    pMensagem = ""
    If Not gfObtRegistro(bdP00MPED, lSql, lrsTabela, pMensagem) Then
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
        lCreate = lCreate & "     NUM_PI   long NOT NULL , "
        lCreate = lCreate & "     IND_IMPRESSAO int NULL, "
        lCreate = lCreate & "     Dsc_LocalExposicao   text  Null ) ;"
        
        If gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call gpGraLog(1, pMensagem)
        End If
        
         lCreate = ""
         lCreate = lCreate & "CREATE  UNIQUE  INDEX Tab_Ped_Espec2_IND03"
         lCreate = lCreate & "       On Tab_Ped_Espec2 (NUM_PI);"
         If gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
             Call gpGraLog(1, pMensagem)
         End If
         
         lCreate = ""
         lCreate = lCreate & "CREATE    INDEX Tab_Ped_Espec2_IND01"
         lCreate = lCreate & "       On Tab_Ped_Espec2 (COD_ORG,COD_RAMO,NUM_PED) ;"
         If gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
             Call gpGraLog(1, pMensagem)
         End If
         
         lCreate = ""
         lCreate = lCreate & "CREATE    INDEX Tab_Ped_Espec2_IND02"
         lCreate = lCreate & "       On Tab_Ped_Espec2 (NUM_APOL,NUM_ENDO);"
         If gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
             Call gpGraLog(1, pMensagem)
         End If
    End If
    
    lSql = "Select top 1 * from Tab_Ped_Espec "
    pMensagem = ""
    If Not gfObtRegistro(bdP00MPED, lSql, lrsTabela, pMensagem) Then
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
        If gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call gpGraLog(1, pMensagem)
        End If
        
        lCreate = ""
        lCreate = lCreate & "CREATE  UNIQUE  INDEX Tab_Ped_Espec_IND03"
        lCreate = lCreate & "       On Tab_Ped_Espec (NUM_PI, NUM_PAG);"
        If gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call gpGraLog(1, pMensagem)
        End If
        
        lCreate = ""
        lCreate = lCreate & "CREATE    INDEX Tab_Ped_Espec_IND01"
        lCreate = lCreate & "       On Tab_Ped_Espec (COD_ORG,COD_RAMO,NUM_PED, NUM_PAG) ;"
        If gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call gpGraLog(1, pMensagem)
        End If
        
        lCreate = ""
        lCreate = lCreate & "CREATE    INDEX Tab_Ped_Espec_IND02"
        lCreate = lCreate & "       On Tab_Ped_Espec (NUM_APOL,NUM_ENDO, NUM_PAG);"
        If gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call gpGraLog(1, pMensagem)
        End If
    End If
    
    lSql = "Select top 1 * from Tab_Ped_Dizer "
    pMensagem = ""
    If Not gfObtRegistro(bdP00MPED, lSql, lrsTabela, pMensagem) Then
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
        If gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call gpGraLog(1, pMensagem)
        End If

        lCreate = ""
        lCreate = lCreate & " CREATE  INDEX  TAB_PED_DIZER_IND01  ON  Tab_Ped_Dizer"
        lCreate = lCreate & "( NUM_PED ,  COD_RAMO ,  COD_ORG )"
        If gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call gpGraLog(1, pMensagem)
        End If
        
        lCreate = ""
        lCreate = lCreate & " CREATE  INDEX  TAB_PED_DIZER_IND02  ON   Tab_Ped_Dizer"
        lCreate = lCreate & "       ( NUM_APOL ,  NUM_ENDO )"
        If gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call gpGraLog(1, pMensagem)
        End If
    End If

    lSql = "Select top 1 * from Tab_Ped_Auto "
    pMensagem = ""
    If Not gfObtRegistro(bdP00MPED, lSql, lrsTabela, pMensagem) Then
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
        If gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call gpGraLog(1, pMensagem)
        End If

        lCreate = ""
        lCreate = lCreate & " CREATE  INDEX  TAB_PED_AUTO_IND01  ON Tab_Ped_Auto"
        lCreate = lCreate & "          ( COD_ORG ,  COD_RAMO ,  NUM_PED )"
        If gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call gpGraLog(1, pMensagem)
        End If

        lCreate = ""
        lCreate = lCreate & " CREATE  INDEX  TAB_PED_AUTO_IND02  ON  Tab_Ped_Auto"
        lCreate = lCreate & "           ( NUM_APOL ,  NUM_ENDO ) "
        If gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call gpGraLog(1, pMensagem)
        End If
    End If
    
    lSql = "Select top 1 * from Tbm_Ramo "
    pMensagem = ""
    If Not gfObtRegistro(bdP00MPED, lSql, lrsTabela, pMensagem) Then
        lCreate = ""
        lCreate = lCreate & "CREATE TABLE  Tbm_Ramo  ("
        lCreate = lCreate & "     COD_RAMO   int  NOT NULL ,"
        lCreate = lCreate & "     DSC_RAMO   char  (30)   NULL ,"
        lCreate = lCreate & "     Tipo_Usuario  char  (1)   NULL )"
        If gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call gpGraLog(1, pMensagem)
        End If
        
        lSql = "Insert into Tbm_Ramo values (112, 'Empresarial','C') "
        If gfExeSQL(bdP00MPED, lSql, pMensagem) <> 0 Then
            Call gpGraLog(1, pMensagem)
        End If
        lSql = "Insert into Tbm_Ramo values (113, 'Residencial','C') "
        If gfExeSQL(bdP00MPED, lSql, pMensagem) <> 0 Then
            Call gpGraLog(1, pMensagem)
        End If
        lSql = "Insert into Tbm_Ramo values (114, 'Condomínio','C') "
        If gfExeSQL(bdP00MPED, lSql, pMensagem) <> 0 Then
            Call gpGraLog(1, pMensagem)
        End If
    End If

    lSql = "Select top 1 Cod_Bem, Cod_Seq from Tab_Ped_Loc_Bem "
    pMensagem = ""
    If Not gfObtRegistro(bdP00MPED, lSql, lrsTabela, pMensagem) Then
        lSql = " alter table Tab_Ped_Loc_Bem add COLUMN Cod_Bem int, "
        lSql = lSql & " Cod_Seq int"
        If gfExeSQL(bdP00MPED, lSql, pMensagem) <> 0 Then
            Call gpGraLog(1, pMensagem)
        End If
        lSql = " Update Tab_Ped_Loc_Bem"
        lSql = lSql & " set Cod_Bem = 999 , Cod_Seq = 1  "
        If gfExeSQL(bdP00MPED, lSql, pMensagem) <> 0 Then
            Call gpGraLog(1, pMensagem)
        End If
        lSql = " Drop Index PrimaryKey on Tab_Ped_Loc_Bem "
        If gfExeSQL(bdP00MPED, lSql, pMensagem) <> 0 Then
            Call gpGraLog(0, pMensagem)
        End If
        lSql = " Create Unique Index PrimaryKey on Tab_Ped_Loc_Bem "
        lSql = lSql & " (Nosso_Numero, Num_Item,  Cod_Bem, Cod_Seq, Dsc_Bem )"
        If gfExeSQL(bdP00MPED, lSql, pMensagem) <> 0 Then
            Call gpGraLog(1, pMensagem)
        End If
    End If
    
    lSql = "Select top 1 * from Tab_Ped_Num_Copia "
    pMensagem = ""
    If Not gfObtRegistro(bdP00MPED, lSql, lrsTabela, pMensagem) Then
        lCreate = ""
        lCreate = lCreate & "CREATE TABLE Tab_Ped_Num_Copia ("
        lCreate = lCreate & "   NOSSO_NUMERO_NEW  char (20) NOT NULL ,"
        lCreate = lCreate & "   NOSSO_NUMERO_OLD  char (20) NOT NULL ,"
        lCreate = lCreate & "   DAT_COPIA  double NOT NULL )"

        If gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call gpGraLog(1, pMensagem)
        End If

        lCreate = ""
        lCreate = lCreate & "  CREATE  INDEX Tab_Ped_Num_Copia_IND01  ON "
        lCreate = lCreate & " Tab_Ped_Num_Copia (NOSSO_NUMERO_NEW) "
        
        If gfExeSQL(bdP00MPED, lCreate, pMensagem) <> 0 Then
            Call gpGraLog(1, pMensagem)
        End If
    End If
    
    
'    '2. Abre base de dados 2.
'    If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, bdP00MULTI, pMensagem) = False Then
'        Call gpGraLog(2, pMensagem)
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
'    If Not gfObtRegistro(bdP00MULTI, lSql, lrsTabela, pMensagem) Then
'        Call gpGraLog(2, pMensagem)
'        End
'    End If
'    If lrsTabela.EOF Then
'        lSql = "Insert into Tbm_Taxa_Classe values ("
''Ficha 1000305
'        lSql = lSql & "113,1,3,0,3,2,5,5.00000,5.00000,5.00000,922,0,0 ) "
'        If gfExeSQL(bdP00MULTI, lSql, pMensagem) <> 0 Then
'            Call gpGraLog(1, pMensagem)
'        End If
'    End If
'    lrsTabela.Close
    
    Call gpFecha2(bdP00MULTI)
    Call gpFecha2(bdP00MPED)
    
    mfCriaTabelas = True
CriaTabelas_Err:
End Function


Private Sub mpCarregaCorretor(ByRef pCod_Corretor As Long)
    'Função: Busca o corretor na arquivo P2842400.EXE

    'Parâmetro de saída.....pCod_Corretor......Código do corretor.
    
Dim lArquivo        As String           'Endereço do arquivo.
Dim lNumArquivo     As Integer          'Utilizado com FreeFile.
Dim lRegistro       As String           'Registro.

    '1. Inicio a variável.
    pCod_Corretor = 0
    
    '2. Busco o dados no arquivo P2842400.EXE
    pCod_Corretor = 0
    lArquivo = gAppPath & "P2842400.EXE"
    If gfPreenchido(Dir(lArquivo)) = False Then
        lArquivo = gAppPath & "P2842401.EXE"
        If gfPreenchido(Dir(lArquivo)) = False Then
            Exit Sub
        End If
    End If
    lNumArquivo = FreeFile
    On Error Resume Next
    Open lArquivo For Input As #lNumArquivo
    If Err = 0 Then
        Line Input #lNumArquivo, lRegistro
        If gfPreenchido(lRegistro) Then
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
    'Função: formata CNPJ.

    Dim lAuxiliar   As String   'Auxiliar.
    Dim lI          As Byte     'Utilizado com For...Next.

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
    'Função: formata CPF.

    Dim lAuxiliar   As String   'Auxiliar.
    Dim lI          As Byte     'Utilizado com For...Next.

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
                               ByRef pMensagem As String, Optional pCod_Ramo As Integer) As Boolean
                               
    'Função: grava registro de pedido novo.

    'Parâmetros de entrada...pBasDados........Base de dados.
    '                        pNosNumero.......Nosso número.
    '                        pNumPedido.......Número do pedido.
    '                        pDatVigInicio....Data de vigência: início.
    '                        pDatVigTermino...Data de vigência: término.
    'Parâmetro de saída......pMensagem........Mensagem.

Dim lCodCorretor    As Long         'Código do corretor.
Dim lInsert         As String       'Comando INSERT.
Dim lRegistro       As String       'Registro.
Dim lObjPed         As clsA46V702A  'Objeto da classe clsA46V702A (TAB_PED)
Dim lObjstcPed      As stcA46V702B  'Objeto da classe de estrutura stcA46V702B (TAB_PED)
Dim lobjCtrlEmis    As clsA46V712A  'Objeto da classe de estrutura stcA46V712B (TAB_CTRL_EMIS)
Dim lObjPedEnd      As clsA46V707A  'Objeto da classe de estrutura stcA46V707B (TAB_PED_ENDER)
Dim lobjCorr        As clsA46V705A  'Objeto da classe clsA46V705A (TAB_PED_CORR)
Dim lObjstcCorr     As stcA46V705B  'Objeto da classe de estrutura stcA46V705B (TAB_PED_CORR)
Dim lobjJuros       As clsA46V711A  'Objeto da classe clsA46V711A (TAB_PED_JUROS)
Dim lQtdRegistro    As Long         'Quantidade de registros.
Dim lObjPedAuto     As clsA46V719A  'Objeto da classe de estrutura stcA46V719B (TAB_PED_AUTO)
Dim lobjEspec       As clsA46V726A  'Objeto da classe de estrutura clsA46V726A (TAB_PED_ESPEC)
Dim lPropTabela     As String

    On Error GoTo ControleErro
    gfGraNovPedido = False
    Call gfObtSYASMULTINI(lPropTabela)

    '1. Obtém código do corretor.
    lCodCorretor = gCod_Corretor
    
    '2. Instancio as classes da tabela TAB_PED
    Set lObjPed = New clsA46V702A
    Set lObjstcPed = New stcA46V702B
    Set lObjPed.ConexaoBD = pBasDados
    
    'Carrego o registro padrão.
    If pCod_Ramo <> 0 Then
        lObjstcPed.Nosso_Numero = "PADRAO" & gNomUsuario6 & pCod_Ramo
    Else
        lObjstcPed.Nosso_Numero = "PADRAO" & gNomUsuario6
    End If
    If Not lObjPed.mfCarregarColecao(lObjstcPed, lPropTabela, "", True) Then
        pMensagem = lObjPed.UltimoErro
        GoTo ControleErro
    End If
    
    If lObjPed.Colecao.Count > 0 Then
        Set lObjstcPed = lObjPed.Colecao.Item(1)
    Else
        Call mfVerRegPadrao(pMensagem)
    End If
    
    If Not IsNumeric(pCod_Ramo) Then pCod_Ramo = 0
    
    'Carrego as informações
    With lObjstcPed
        .Nosso_Numero = pNosNumero
        .Num_Ped = pNumPedido
        .Cod_Ramo = pCod_Ramo
        .Dat_Ini_Vig = pDatVigInicio
        .Dat_Ter_Vig = pDatVigTermino
        .Cod_Tip_Doc = e_TipDoc_Cotacao
        .Cod_Produto = pCodProduto
        .Cod_Protocolo = pCodProtocolo
        .Cod_Sequencia = pCodSequencia
        .Cod_User_Liber = gNomUsuario6
        .Cod_Dizer = 0
        If gNum_Pi = 0 Then
            .Num_PI = Right(Year(Date), 2) & Format(pNumPedido, "00000000")
        Else
            .Num_PI = gNum_Pi
        End If
        .Cod_Reg = 1
        .Tip_Cust_Emis = 1
        .Val_Cust_Emis = gProCusApolice
        .Val_Parc_Min = gProPreMinParcela
        If gTipUsuario = e_Tipo_Usu_CorretorYasuda Then
            .Cod_Segurado = gCodSeguradoYas
        End If
        If .Cod_Estip = 0 Then .Cod_Estip = 1
        .Cod_Iden_Praz = e_TipoPrazo_ProRata
        .Qtd_Parc = 10
        'Fichas 1200168 e 1200169 Inclusão dos campos de listagem Renda e Faixa de Faturamento Pessoa física e jurídica
        .Val_RendaFatu = -1
    End With
    '3. Chamo a rotina de inclusão
    If Not lObjPed.mfIncluir(lObjstcPed, lPropTabela, "") Then
        pMensagem = lObjPed.UltimoErro
        GoTo ControleErro
    End If
    
    'CTRL_EMIS
    Set lobjCtrlEmis = New clsA46V712A
    Set lobjCtrlEmis.ConexaoBD = lObjPed.ConexaoBD
    lObjstcPed.objstcCtrlEmis.Nosso_Numero = pNosNumero
    lObjstcPed.objstcCtrlEmis.Cod_Ramo = pCod_Ramo
    If Not lobjCtrlEmis.mfIncluir(lObjstcPed.objstcCtrlEmis, lPropTabela, "") Then
        pMensagem = lobjCtrlEmis.UltimoErro
        GoTo ControleErro
    End If
    
    'PED_ENDERECO
    Set lObjPedEnd = New clsA46V707A
    Set lObjPedEnd.ConexaoBD = lObjPed.ConexaoBD
    lObjstcPed.objstcPedEndereco.Nosso_Numero = pNosNumero
    lObjstcPed.objstcPedEndereco.Cod_Ramo = pCod_Ramo
    If Not lObjPedEnd.mfIncluir(lObjstcPed.objstcPedEndereco, lPropTabela, "") Then
        pMensagem = lObjPedEnd.UltimoErro
        GoTo ControleErro
    End If
    
    'PED_AUTO
    Set lObjPedAuto = New clsA46V719A
    Set lObjPedAuto.ConexaoBD = lObjPed.ConexaoBD
    lObjstcPed.objstcPedAuto.Nosso_Numero = pNosNumero
    lObjstcPed.objstcPedAuto.Cod_Ramo = pCod_Ramo
    If Not lObjPedAuto.mfIncluir(lObjstcPed.objstcPedAuto, lPropTabela, "") Then
        pMensagem = lObjPedEnd.UltimoErro
        GoTo ControleErro
    End If
    
    'Juros
    'Set lobjJuros = New clsA46V711A
    'Set lobjJuros.ConexaoBD = lObjPed.ConexaoBD
    'If Not lobjJuros.mfCopiarPadraoParaNovo(pNosNumero) Then
    '    pMensagem = lobjJuros.UltimoErro
    '    GoTo ControleErro
    'End If
    
    'PED_CORR
    Set lobjCorr = New clsA46V705A
    Set lobjCorr.ConexaoBD = lObjPed.ConexaoBD
    For Each lObjstcCorr In lObjstcPed.ColecaoPedCorr
        lObjstcCorr.Nosso_Numero = pNosNumero
        lObjstcCorr.Cod_Org = lObjstcPed.Cod_Org
        lObjstcCorr.Cod_Ramo = pCod_Ramo
        lObjstcCorr.Num_Ped = pNumPedido
'        lObjstcCorr.Cod_Corr = lCodCorretor
'        lObjstcCorr.Num_Ordem_Corr = 1
        lObjstcCorr.Num_PI = lObjstcPed.Num_PI
        Set lObjstcCorr.ObjPai = lObjstcPed
        If Not lobjCorr.mfIncluir(lObjstcCorr, lPropTabela, "") Then
            pMensagem = lobjCorr.UltimoErro
            GoTo ControleErro
        End If
    Next
    
    'TAB_PED_ESPEC
    Dim lobjstcEspc As stcA46V726B
    Set lobjEspec = New clsA46V726A
    Set lobjEspec.ConexaoBD = lObjPed.ConexaoBD
    
    For Each lobjstcEspc In lObjstcPed.ColecaoTextoEspec
        lobjstcEspc.Nosso_Numero = pNosNumero
        lobjstcEspc.Cod_Org = lObjstcPed.Cod_Org
        lobjstcEspc.Cod_Ramo = pCod_Ramo
        lobjstcEspc.Num_Ped = pNumPedido
        lobjstcEspc.Num_PI = lObjstcPed.Num_PI
        Set lobjstcEspc.ObjPai = lObjstcPed
        If Not lobjEspec.mfIncluir(lobjstcEspc, lPropTabela, "") Then
            pMensagem = lobjEspec.UltimoErro
            GoTo ControleErro
        End If
    Next
    
   'TAB_PED_ESPEC2
   'Set lobjEspec = New clsA46V721A
   'Set lobjEspec.ConexaoBD = lObjPed.ConexaoBD
   'lObjstcPed.objstcEspec.Nosso_Numero = pNosNumero
   'If Not lobjEspec.mfIncluir(lObjstcPed.objstcEspec) Then
   '    pMensagem = lobjEspec.UltimoErro
   '    GoTo ControleErro
   'End If
    
    gfGraNovPedido = True
    Set pBasDados = lObjPed.ConexaoBD
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
    Call gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gpGraLog(0, "gfGraNovPedido ERRO - " & Err & " " & Err.Description & " " & pMensagem)
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
                               ByRef pNosNumero As String, pIndSinistro As e_SimNao, _
                               ByRef pMensagem As String) As Boolean
    'Função: grava registros de cotação/proposta de renovação Yasuda.

    'Parâmetros de entrada...pNumApolice...Número da apólice.
    '                        pIteApolice...Número do item da apólice.
    '                        pNosNumero....Nosso número.
    'Parâmetro de saída......pMensagem.....Mensagem.

    Dim bdP00APOL       As ADODB.Connection 'Base de dados: P00Apol.
    Dim lCodProduto     As Integer          'Código do produto.
    Dim lCodProtocolo   As Long             'Código do protocolo.
    Dim lCodSequencia   As Long             'Código da seqüência.
    Dim lDescTarifa     As String           'Descrição da tarifa.
    Dim lNosNumero      As String           'Nosso número.
    Dim lNumPedido      As String           'Número do pedido.
    Dim lQuaIteComErro  As Integer          'Quantidade de itens com erro.
    Dim lQuaIteSemErro  As Integer          'Quantidade de itens sem erro.
    Dim lVigDatIni      As Long             'Data de vigência: início.
    Dim lVigDatTer      As Long             'Data de vigência: término.
    Dim rsP00APOL       As ADODB.Recordset  'Registro: P00Apol.
    Dim lNumEndo        As Long
    Dim lRegistro       As String
    Dim lNomeTab        As String
    Dim lNumItem        As Integer
    Dim lAchouItem      As Boolean
    Dim rdsTemp         As ADODB.Recordset  '1100177 - Renovação Apolice LMI
    Dim lNumCotac_LMI   As String           '1100177 - Renovação Apolice LMI
    Dim lSelect         As String
    
    On Error GoTo ControleErro
    gfGraRenYasuda = False
    
    If gPrograma = "D046O003" Then
        lNomeTab = "Syasportal.dbo.TAB_APOL"
    Else
        If gBaseLocal = True Then
            lNomeTab = "TAB_APOL"
        Else
            lNomeTab = "GED.DBO.TAB_PED"
        End If
    End If
    
    lNumCotac_LMI = "" '1100177 - Renovação Apolice LMI
    
BuscaApolice:
    '1. Obtém registro da apólice.
    'a) Monta comando SELECT.
    lSelect = "SELECT distinct "
    lSelect = lSelect & lNomeTab & ".*, NUM_ITEM, COD_TIP_LOGR_RISCO, "
    lSelect = lSelect & " NOM_LOGR_RISCO, "
    If gBaseLocal = True Or gPrograma = "D046O003" Then
        lSelect = lSelect & " Dsc_Compl_Risco , "
    Else
        lSelect = lSelect & " left('N.' + (select cast(NUM_LOGR_RISCO as varchar) from " & lNomeTab & "_Insp where "
        If lNumCotac_LMI = "" Then    '1100177 - Renovação Apolice LMI
            lSelect = lSelect & lNomeTab & "_Insp.Num_PI = " & lNomeTab & "_loc.num_pi and " & lNomeTab & "_Insp.num_ped = "
        Else
            lSelect = lSelect & lNomeTab & "_Insp.NOSSO_NUMERO = " & lNomeTab & "_loc.NOSSO_NUMERO and " & lNomeTab & "_Insp.num_ped = "
        End If
        lSelect = lSelect & lNomeTab & "_loc.num_ped and " & lNomeTab & "_Insp.num_item = " & lNomeTab & "_loc.num_item) +' - '+  rtrim(DSC_COMPL_RISCO) ,30) as Dsc_Compl_Risco,"
    End If
    lSelect = lSelect & " Nom_Bairro_Risco , "
    lSelect = lSelect & lNomeTab & "_LOC.NUM_PI , "  '1100177 - Renovação Apolice LMI
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
    lSelect = lSelect & " " & lNomeTab & "_LOC.NUM_ENDO AS NUM_ENDO2 "
    lSelect = lSelect & " FROM " & lNomeTab & "_LOC INNER JOIN " & lNomeTab & " ON "
    lSelect = lSelect & " " & lNomeTab & "_LOC.NUM_APOL = " & lNomeTab & ".NUM_APOL "
    If (Not gBaseLocal) And (Not gPrograma = "D046O003") Then
        lSelect = lSelect & " AND " & lNomeTab & "_LOC.NUM_ENDO = " & lNomeTab & ".NUM_ENDO "
    End If
    If lNumCotac_LMI = "" Then    '1100177 - Renovação Apolice LMI
        lSelect = lSelect & " Where " & lNomeTab & "_LOC.Num_Apol = " & pNumApolice
    Else
        lSelect = lSelect & " AND   " & lNomeTab & "_LOC.NOSSO_NUMERO = " & lNomeTab & ".NOSSO_NUMERO"
        lSelect = lSelect & " Where " & lNomeTab & ".NOSSO_NUMERO = " & mfFormatarCamposSQL(lNumCotac_LMI)
    End If
    'Ficha 1000782 - Não permitir Renovação LMI para Corretores
    If gTipUsuario = e_Tipo_Usu_Corretor Then
        lSelect = lSelect & " AND " & lNomeTab & "_LOC.NOM_LOGR_RISCO not like '%CONFORME ESPEC%'"
    End If
    lSelect = lSelect & " AND [COD_DIZER] NOT IN (4,10)          "
    If Val(pIteApolice) <> 0 Then
        lSelect = lSelect & " AND NUM_ITEM = " & pIteApolice
    End If
   ' If Not gBaseLocal Then
   '     lSelect = lSelect & " and " & lNomeTab & ".NUM_PI IN (SELECT NUM_PI "
   '     lSelect = lSelect & " From [tab_ctrl_emis]"
   '     lSelect = lSelect & " WHERE  [cod_iden_stat_acao] = 90 ) "
   ' End If
    'lSelect = lSelect & " ORDER BY " & lNomeTab & ".NUM_ENDO desc ,NUM_ITEM"
    lSelect = lSelect & " ORDER BY NUM_ITEM, " & lNomeTab & ".DAT_EMIS desc"
    
    'b) Abre arquivo.
    If gfAbrBasDados("P00APOL", gTipP00APOL, gEndP00APOL, bdP00APOL, pMensagem) = False Then
        Exit Function
    End If
    'c) Obtém registro.
    If gfObtRegistro(bdP00APOL, lSelect, rsP00APOL, pMensagem) = False Then
        Call gpFecha2(bdP00APOL)
        Exit Function
    End If
    lAchouItem = True
    'd) Não encontrou registro de apólice.
    If rsP00APOL.EOF = True Then
        '1001106 - Renovação Roubo
        lSelect = "SELECT * FROM " & lNomeTab
        lSelect = lSelect & " Where Num_Apol = " & pNumApolice
        lSelect = lSelect & " AND NUM_ENDO = 0"
        If gfObtRegistro(bdP00APOL, lSelect, rsP00APOL, pMensagem) = False Then
            Call gpFecha2(bdP00APOL)
            Exit Function
        End If
        If rsP00APOL.EOF = True Then
            '===1) Se não encontrou registro, verifica se tem apólice.
            If gfPreenchido(pIteApolice) Then
                lSelect = "SELECT * FROM " & lNomeTab & "_LOC "
                lSelect = lSelect & " Where Num_Apol = " & pNumApolice
                lSelect = lSelect & " order by NUM_ENDO desc "
                If gfObtRegistro(bdP00APOL, lSelect, rsP00APOL, pMensagem) = False Then
                    Call gpFecha2(bdP00APOL)
                    Exit Function
                End If
            End If
            If rsP00APOL.EOF = True Then
Verifica:
                pMensagem = "FNL0012 - Apólice " & pNumApolice & " não localizada."
            Else
                pMensagem = "FNL0013 - Item  " & pIteApolice & " não localizado na apólice " & _
                            pNumApolice & "."
            End If
            Call gpFecha3(rsP00APOL, bdP00APOL)
            Exit Function
        Else
            If rsP00APOL("cod_ramo") <> 150 And rsP00APOL("cod_ramo") <> 710 Then
                GoTo Verifica
            End If
            lAchouItem = False
        End If
    End If

    '1100177 - Renovação Apolice LMI
    If (gTipUsuario = e_Tipo_Usu_Yasuda) And (gBaseLocal = False) And (lNumCotac_LMI = "") Then
        If Left(rsP00APOL("NOM_LOGR_RISCO"), 14) = "CONFORME ESPEC" Then
            rsP00APOL.Filter = " Num_Endo = 0"
            lSelect = "Select NUM_COTAC_LMI FROM RAMOSDIVERSOS.DBO.TAB_PED"
            lSelect = lSelect & " Where NUM_PI = " & rsP00APOL("NUM_PI")
            If gfObtRegistro(bdP00APOL, lSelect, rdsTemp, pMensagem) = False Then
                Call gpFecha2(bdP00APOL)
                Exit Function
            End If
            If rdsTemp.EOF Then
                GoTo Verifica
            ElseIf Trim(gfTratarNulo(rdsTemp("NUM_COTAC_LMI"), TpoString)) <> "" Then
                lNumCotac_LMI = Trim(gfTratarNulo(rdsTemp("NUM_COTAC_LMI"), TpoString))
                lNomeTab = "RAMOSDIVERSOS.DBO.TAB_PED"
                GoTo BuscaApolice
            End If
        End If
    End If
    
    '3. Posiciona nosso número.
    If Not gfPreenchido(pNosNumero) Then
        If gfCalNosNumero(lNosNumero, pMensagem) = False Then
            Call gpFecha3(rsP00APOL, bdP00APOL)
            Exit Function
        End If
        pNosNumero = lNosNumero
    Else
        lNosNumero = pNosNumero
    End If

    '4. Posiciona número do pedido.
    If gfCalNumPedido(lNumPedido, pMensagem) = False Then
        Call gpFecha3(rsP00APOL, bdP00APOL)
        Exit Function
    End If
    gPrefixo = "Cotação " & lNumPedido & ". "

    '5. Posiciona vigência.
    rsP00APOL.MoveFirst
    
    lVigDatIni = rsP00APOL!Dat_Ter_Vig
    lVigDatTer = lVigDatIni + 10000
    
    '6. Posiciona produto que atenda à vigência.
    If gfBusProduto(Format(Date, "YYYYMMDD"), lCodProduto, lCodProtocolo, lCodSequencia, lDescTarifa, pMensagem, rsP00APOL("Cod_Ramo")) = _
       False Then
        Call gpFecha3(rsP00APOL, bdP00APOL)
        Exit Function
    End If
    If lCodProduto = 0 Then
        pMensagem = "FGE0088 - Não há produto (tarifa) disponível para cálculo da renovação da " & _
                    "apólice " & pNumApolice & ", de acordo com a data de início de vigência (" & _
                    gfForData(2, lVigDatIni) & ")." & vbLf & vbLf & "Aguardar novo produto (tarifa)." & _
                    vbLf & vbLf & "Caso seja necessário, o cálculo deve ser feito pela Yasuda."
        Exit Function
    End If
    'Posicionar o mesmo plano
    
'sandra 900218
    gCod_Plano = rsP00APOL!Cod_Estip
    If gCod_Plano >= 1001 And gCod_Plano <= 1010 Then
    'De/Para dos planos da Mercedes
        Select Case gCod_Plano
            Case 1001
                gCod_Plano = 2021
            Case 1002
                gCod_Plano = 2022
            Case 1003
                gCod_Plano = 2023
            Case 1004
                gCod_Plano = 2024
            Case 1005
                gCod_Plano = 2025
            Case 1006
                gCod_Plano = 2026
            Case 1007
                gCod_Plano = 2027
            Case 1008
                gCod_Plano = 2028
            Case 1009
                gCod_Plano = 2029
            Case 1010
                gCod_Plano = 2030
        End Select
    
        gProCusApolice = 20
        gProPreMinParcela = 0
    End If
    
    '7. Grava registro de renovação Yasuda.
    If gCod_Ramo = e_CodRamo_Riscos_diversos Then
        If mfGraRenComum(rsP00APOL, lNosNumero, CLng(lNumPedido), lVigDatIni, lVigDatTer, lCodProduto, _
           CInt(lCodProtocolo), CInt(lCodSequencia), gCod_Dizer, bdP00APOL, lNumCotac_LMI, pNumApolice, pMensagem) = False Then
            Call gpFecha3(rsP00APOL, bdP00APOL)
            Exit Function
        End If
    Else
        If mfGraRenComum(rsP00APOL, lNosNumero, CLng(lNumPedido), lVigDatIni, lVigDatTer, lCodProduto, _
           CInt(lCodProtocolo), CInt(lCodSequencia), 0, bdP00APOL, lNumCotac_LMI, pNumApolice, pMensagem) = False Then
            Call gpFecha3(rsP00APOL, bdP00APOL)
            Exit Function
        End If
    End If
    gCodProduto = lCodProduto

    '8. Grava registros de itens, coberturas e descontos.
    rsP00APOL.MoveLast      'Posiciona quantidade de registros.
    rsP00APOL.MoveFirst
    lNumEndo = 0
    lNumItem = 0
    While Not rsP00APOL.EOF
        '1001106 - Renovação Roubo
        If (rsP00APOL("COD_RAMO") = e_CodRamo_Roubos Or rsP00APOL("COD_RAMO") = e_CodRamo_Riscos_diversos) And lAchouItem = False Then
            If Not mfGraRenItem_Roubo(rsP00APOL, lNosNumero, bdP00APOL, pMensagem, pIndSinistro, CLng(lNumPedido), 0) Then
                Call gpGraLog(1, pMensagem)
                Exit Function
            End If
        Else
            If lNumItem <> rsP00APOL("Num_Item") Then
                If Not mfGraRenItem(rsP00APOL, lNosNumero, bdP00APOL, _
                        pMensagem, pIndSinistro, CLng(lNumPedido), gCod_Dizer, lNumCotac_LMI, pNumApolice) Then
                    Call gpGraLog(1, pMensagem)
                    Exit Function
                End If
                lNumItem = rsP00APOL("Num_Item")
            End If
        End If
        rsP00APOL.MoveNext
    Wend
    Call gpFecha3(rsP00APOL, bdP00APOL)

    '9. Grava registro de comunicação.
    If gfGraRegComunicacao(lNosNumero, pMensagem) = False Then
        Call gpGraLog(1, pMensagem)
        Exit Function
    End If

    gfGraRenYasuda = True
    Exit Function
ControleErro:
    Call gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gpGraLog(0, "gfGraRenYasuda ERRO - " & Err & " " & Err.Description & " " & pMensagem)
End Function

Public Function gfGraEndossoYasuda(ByVal pNumApolice As String, ByVal pIteApolice As String, _
                               ByRef pNosNumero As String, pIndSinistro As e_SimNao, _
                               ByVal pCod_Dizer As Integer, ByRef pMensagem As String) As Boolean
    'Função: grava registros de cotação/proposta de Endosso Yasuda.

    'Parâmetros de entrada...pNumApolice...Número da apólice.
    '                        pIteApolice...Número do item da apólice.
    '                        pNosNumero....Nosso número.
    'Parâmetro de saída......pMensagem.....Mensagem.

    Dim bdP00APOL       As ADODB.Connection 'Base de dados: P00Apol.
    Dim bdP00MPED       As ADODB.Connection 'Base de dados: P00MPED.
    Dim lCodProduto     As Integer          'Código do produto.
    Dim lCodProtocolo   As Long             'Código do protocolo.
    Dim lCodSequencia   As Long             'Código da seqüência.
    Dim lDescTarifa     As String           'Descrição da tarifa.
    Dim lNosNumero      As String           'Nosso número.
    Dim lNumPedido      As String           'Número do pedido.
    Dim lQuaIteComErro  As Integer          'Quantidade de itens com erro.
    Dim lQuaIteSemErro  As Integer          'Quantidade de itens sem erro.
    Dim lVigDatIni      As Long             'Data de vigência: início.
    Dim lVigDatTer      As Long             'Data de vigência: término.
    Dim rsP00APOL       As ADODB.Recordset  'Registro: P00Apol.
    Dim lNumEndo        As Long
    Dim lRegistro       As String
    Dim lNomeTab        As String
    Dim lNumItem        As Integer
    Dim lSelect         As String
    
    On Error GoTo ControleErro
    gfGraEndossoYasuda = False
    
    If gBaseLocal = True Then
        lNomeTab = "TAB_APOL"
    Else
        lNomeTab = "GED.DBO.TAB_PED"
    End If
    
    '1. Obtém registro da apólice.
    'a) Monta comando SELECT.
    lSelect = "SELECT distinct "
    lSelect = lSelect & " " & lNomeTab & ".*, NUM_ITEM, COD_TIP_LOGR_RISCO, "
    lSelect = lSelect & " NOM_LOGR_RISCO, DSC_COMPL_RISCO, NOM_BAIRRO_RISCO, "
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
    lSelect = lSelect & " " & lNomeTab & "_LOC.COD_PRODUTO, COD_REGIAO, COD_IDEN_TARIF"
    lSelect = lSelect & " FROM " & lNomeTab & "_LOC INNER JOIN " & lNomeTab & " ON "
    lSelect = lSelect & " " & lNomeTab & "_LOC.NUM_APOL = " & lNomeTab & ".NUM_APOL "
    lSelect = lSelect & " Where " & lNomeTab & "_LOC.Num_Apol = " & pNumApolice
    If Val(pIteApolice) <> 0 Then
        lSelect = lSelect & " AND NUM_ITEM = " & pIteApolice
    End If
    lSelect = lSelect & " ORDER BY " & lNomeTab & ".NUM_ENDO desc ,NUM_ITEM"
    
    'b) Abre arquivo.
    If gfAbrBasDados("P00APOL", gTipP00APOL, gEndP00APOL, bdP00APOL, pMensagem) = False Then
        Exit Function
    End If
    'c) Obtém registro.
    If gfObtRegistro(bdP00APOL, lSelect, rsP00APOL, pMensagem) = False Then
        Call gpFecha2(bdP00APOL)
        Exit Function
    End If
    
    'd) Não encontrou registro de apólice.
    If rsP00APOL.EOF = True Then
        '===1) Se não encontrou registro, verifica se tem apólice.
        If gfPreenchido(pIteApolice) Then
            lSelect = "SELECT * FROM " & lNomeTab & "_LOC "
            lSelect = lSelect & " Where Num_Apol = " & pNumApolice
            lSelect = lSelect & " order by NUM_ENDO desc "
            If gfObtRegistro(bdP00APOL, lSelect, rsP00APOL, pMensagem) = False Then
                Call gpFecha2(bdP00APOL)
                Exit Function
            End If
        End If
        If rsP00APOL.EOF = True Then
            pMensagem = "FNL0012 - Apólice " & pNumApolice & " não localizada."
        Else
            pMensagem = "FNL0013 - Item  " & pIteApolice & " não localizado na apólice " & _
                        pNumApolice & "."
        End If
        Call gpFecha3(rsP00APOL, bdP00APOL)
        Exit Function
    End If

    '3. Posiciona nosso número.
    If gfPreenchido(pNosNumero) Then
        lNosNumero = pNosNumero
    Else
        If gfCalNosNumero(lNosNumero, pMensagem) = False Then
            Call gpFecha3(rsP00APOL, bdP00APOL)
            Exit Function
        End If
        pNosNumero = lNosNumero
    End If

    '4. Posiciona número do pedido.
    If gfCalNumPedido(lNumPedido, pMensagem) = False Then
        Call gpFecha3(rsP00APOL, bdP00APOL)
        Exit Function
    End If
    gPrefixo = "Cotação " & lNumPedido & ". "

    '5. Posiciona vigência.
    rsP00APOL.MoveFirst

    lVigDatIni = gHoje
    lVigDatTer = rsP00APOL!Dat_Ter_Vig
    
    '6. Posiciona produto que atenda à vigência.
    If gfBusProduto(lVigDatIni, lCodProduto, lCodProtocolo, lCodSequencia, lDescTarifa, pMensagem, rsP00APOL("Cod_Ramo")) = _
       False Then
        Call gpFecha3(rsP00APOL, bdP00APOL)
        Exit Function
    End If
    If lCodProduto = 0 Then
        pMensagem = "FGE0088 - Não há produto (tarifa) disponível para cálculo da renovação da " & _
                    "apólice " & pNumApolice & ", de acordo com a data de início de vigência (" & _
                    gfForData(2, lVigDatIni) & ")." & vbLf & vbLf & "Aguardar novo produto (tarifa)." & _
                    vbLf & vbLf & "Caso seja necessário, o cálculo deve ser feito pela Yasuda."
        Exit Function
    End If
    'Posicionar o mesmo plano
    gCod_Plano = rsP00APOL!Cod_Estip
    
    '7. Grava registro de renovação Yasuda.
    If mfGraRenComum(rsP00APOL, lNosNumero, CLng(lNumPedido), lVigDatIni, lVigDatTer, lCodProduto, _
       CInt(lCodProtocolo), CInt(lCodSequencia), pCod_Dizer, bdP00APOL, "", pNumApolice, pMensagem) = False Then
        Call gpFecha3(rsP00APOL, bdP00APOL)
        Exit Function
    End If

    '8. Grava registros de itens, coberturas e descontos.
    'Caso seja encosso de inclusão de local de risco não inclui item
    If pCod_Dizer <> 1 Then
        rsP00APOL.MoveLast      'Posiciona quantidade de registros.
        rsP00APOL.MoveFirst
        lNumEndo = 0
        While Not rsP00APOL.EOF
            If Not mfGraRenItem(rsP00APOL, lNosNumero, bdP00APOL, _
                    pMensagem, pIndSinistro, CLng(lNumPedido), pCod_Dizer, "", pNumApolice) Then
                Call gpGraLog(1, pMensagem)
                Exit Function
            End If
            rsP00APOL.MoveNext
        Wend
    Else
        'Acha último Item
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
        If gfAbrBasDados("P00MPED", gTipP00MPED, gEndP00MPED, bdP00MPED, pMensagem) = False Then
            Exit Function
        End If
        'Grava novo registro de item.
        If gfGraItem(bdP00MPED, lNosNumero, lNumPedido, lNumItem, pMensagem, 100) = False Then
            Exit Function
        End If
        Call gpFecha2(bdP00MPED)
    End If
    Call gpFecha3(rsP00APOL, bdP00APOL)

    '9. Grava registro de comunicação.
    If gfGraRegComunicacao(lNosNumero, pMensagem) = False Then
        Call gpGraLog(1, pMensagem)
        Exit Function
    End If

    gfGraEndossoYasuda = True
    Exit Function
ControleErro:
    Call gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gpGraLog(0, "gfGraRenYasuda ERRO - " & Err & " " & Err.Description & " " & pMensagem)
End Function


Private Function mfGraRenComum(pRstApol As ADODB.Recordset, pNosNumero As String, pNumPedido As Long, _
                    pVigDatIni As Long, pVigDatTer As Long, pCodProduto As Integer, _
                    pCodProtocolo As Integer, pCodSequencia As Integer, pCod_Dizer As Integer, _
                    pBdP00APOL As ADODB.Connection, pNumCotac_LMI As String, ByVal pNumApolice As String, pMensagem As String) As Boolean
    'Função: Grava os registros de comum com os dados da apólice (Renovação Yasuda).
    'Parâmetros de entrada :pNum_Apol........Número da apólice que está sendo renovada.
    '                       pNosNumero.......Nosso número.
    '                       pNumPedido.......Número do pedido.
    '                       pVigDatIni.......Data de início da vigência.
    '                       pVigDatTer.......Data de término da vigência.
    '                       pCodProduto......Código do produto.
    '                       pCodProtocolo....Código do protocolo.
    '                       pCodSequencia....Código da sequencia.
    '                       pBdP00APOL.......Arquivo de conexão
    'Parâmetro de saída:    pMensagem........Mensagem de erro.

    Dim lSql            As String                   'Consulta SQL.
    Dim rsAPOL          As ADODB.Recordset          'Registro.
    Dim lobjPedido      As clsA46V702A              'Objeto da classe clsA46V702A (TAB_PED).
    Dim lobjstcPedido   As stcA46V702B              'Objeto da classe de estrutura stcA46V702B (TAB_PED).
    Dim lobjCorr        As clsA46V705A              'Objeto da classe clsA46V705A (TAB_PED_CORR).
    Dim lObjstcCorr     As stcA46V705B              'Objeto da classe de estrutura stcA46V705B (TAB_PED_CORR).
    Dim lobjCoss        As clsA46V706A              'Objeto da classe clsA46V705A (TAB_PED_CORR).
    Dim lobjstcCoss     As stcA46V706B              'Objeto da classe de estrutura stcA46V705B (TAB_PED_CORR).
    Dim lObjEnd         As clsA46V707A              'Objeto da classe clsA46V707A (TAB_PED_ENDER).
    Dim lObjstcEnd      As stcA46V707B              'Objeto da classe de estrutura stcA46V707B (TAB_PED_ENDER).
    Dim lobjCtrlEmis    As clsA46V712A              'Objeto da classe clsA46V712A (TAB_CTRL_EMIS).
    Dim lobjstcCtrlEmis As stcA46V712B              'Objeto da classe de estrutura stcA46V712B (TAB_CTRL_EMIS).
    Dim lobjJuros       As clsA46V711A              'Objeto da classe clsA46V711A (TAB_PED_JUROS).
    Dim lobjAuto        As clsA46V719A              'Objeto da classe clsA46V707A (TAB_PED_AUTO).
    Dim lobjEspec       As clsA46V721A              'Objeto da classe clsA46V721A (TAB_PED_ESPEC2).
    Dim lobjstcEspec    As stcA46V721B              'Objeto da classe de estrutura stcA46V721B (TAB_PED_ESPEC).
    Dim clsA99V001      As clsA99V001
    Dim lNomeTab        As String
    Dim lDigito         As String
    Dim lAuxiliar       As String           'Auxiliar.
    Dim lPropTabela     As String
    
    
    On Error GoTo ControleErro
    Call gfObtSYASMULTINI(lPropTabela)
    
    If Not pRstApol.EOF Then
        Set lobjPedido = New clsA46V702A
        Set lobjstcPedido = New stcA46V702B
        With lobjstcPedido
            .Nosso_Numero = pNosNumero
            .Num_Ped = pNumPedido
            .Cod_Ramo = gfTratarNulo(pRstApol("Cod_Ramo"), tpoNumerico)
            .Cod_Org = gfTratarNulo(pRstApol("Cod_Org"), tpoNumerico)
            .Num_Apol_An = pRstApol("Num_Apol")
            .Cod_Unid_Prod = gfTratarNulo(pRstApol("Cod_Unid_Prod"), tpoNumerico)
            
            'ficha - 801312 (ajustado em 23/04)
            '.Cod_Prod = gfTratarNulo(pRstApol("Cod_Prod"), tpoNumerico)
            .Cod_Prod = 0
            .Cod_Segurado = gfTratarNulo(pRstApol("Cod_Segurado"), tpoNumerico)
            .Nom_Segurado = gfTratarNulo(pRstApol("Nom_Segurado"), TpoString)
            .Dsc_Compl_Segurado = gfTratarNulo(pRstApol("Dsc_Compl_Segurado"), TpoString)
            .Num_Cgc_Cpf = gfTratarNulo(pRstApol("Num_Cgc_Cpf"), tpoNumerico)
            .Cod_Tip_Logr = gfTratarNulo(pRstApol("Cod_Tip_Logr"), TpoString)
            .Nom_Logr = gfTratarNulo(pRstApol("Nom_Logr"), TpoString)
            .Dsc_Compl_Logr = gfTratarNulo(pRstApol("Dsc_Compl_Logr"), TpoString)
            .Nom_Bairro = gfTratarNulo(pRstApol("Nom_Bairro"), TpoString)
            .Nom_Cid = gfTratarNulo(pRstApol("Nom_Cid"), TpoString)
            .Sig_UF = gfTratarNulo(pRstApol("Sig_UF"), TpoString)
            .Num_CEP = gfTratarNulo(pRstApol("Num_Cep"), tpoNumerico)
            .Num_Fone = gfTratarNulo(pRstApol("Num_Fone"), tpoNumerico)
            .Dat_Ini_Vig = pVigDatIni
            .Dat_Ter_Vig = pVigDatTer
            .Cod_Iden_Praz = e_TipoPrazo_ProRata
            .Cod_User_Liber = gNomUsuario6
            .Dat_Liber = 0
            .Hor_Liber = 0
            .Cod_Sit_Ped = e_SitPedido_SemCalculo
            .Cod_Tip_Doc = e_TipDoc_Cotacao
            If pCod_Dizer = 0 Then
                .Cod_Produto = pCodProduto
            Else
                .Cod_Produto = gfTratarNulo(pRstApol("COD_IDEN_TARIF"), tpoNumerico)
            End If
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
            
            .Cod_Estip = gfTratarNulo(pRstApol("COD_ESTIP"), tpoNumerico)
            
            If gNum_Pi = 0 Then
                .Num_PI = Right(Year(Date), 2) & Format(pNumPedido, "00000000")
            Else
                .Num_PI = gNum_Pi
            End If
            If pCod_Dizer = 0 Then
                .Val_Cust_Emis = gProCusApolice
            Else
                .Val_Cust_Emis = gProCusEndosso
            End If
            .Tip_Cust_Emis = 1
            .Val_Parc_Min = gProPreMinParcela
            .Cod_Bco_Cobr = gfTratarNulo(pRstApol("Cod_Bco_Cobr"), tpoNumerico)
            .Cod_Agen_Bas_Cor = gfTratarNulo(pRstApol("Cod_Agen_Bas_Cor"), tpoNumerico)
            If gTipUsuario = e_Tipo_Usu_Corretor Or gTipUsuario = e_Tipo_Usu_CorretorYasuda Then
                lDigito = 0
                Set clsA99V001 = New clsA99V001
                Select Case .Cod_Bco_Cobr
                    Case 1  'Projeto 11080 - Debito automatico - Banco do Brasil
                        Call clsA99V001.DIG_BANCOBRASIL(Format(gfTratarNulo(pRstApol("Num_Cta_Bas_Cor"), tpoNumerico), "0000000"), lDigito, lAuxiliar)
                    Case 237
                        Call clsA99V001.DIG_BRADESCO(Format(gfTratarNulo(pRstApol("Num_Cta_Bas_Cor"), tpoNumerico), "0000000"), lDigito, lAuxiliar)
                    Case 341
                        Call clsA99V001.DIG_ITAU(Format(.gAgenciaDig, "0000"), Format(gfTratarNulo(pRstApol("Num_Cta_Bas_Cor"), tpoNumerico), "00000"), lDigito, lAuxiliar)
                    Case 347
                        If .gAgenciaDig >= 1000 Then
                            Call clsA99V001.DIG_SUDAMERIS_NOVO(Format(.gAgenciaDig, "0000"), Format(gfTratarNulo(pRstApol("Num_Cta_Bas_Cor"), tpoNumerico), "0000000"), lDigito, lAuxiliar)
                        Else
                            Call clsA99V001.DIG_SUDAMERIS(Format(.gAgenciaDig, "000"), Format(gfTratarNulo(pRstApol("Num_Cta_Bas_Cor"), tpoNumerico), "000000000"), lDigito, lAuxiliar)
                        End If
                    Case 409
                         Call clsA99V001.DIG_UNIBANCO(Format(.gAgenciaDig, "0000"), Format(gfTratarNulo(pRstApol("Num_Cta_Bas_Cor"), tpoNumerico), "000000"), lDigito, lAuxiliar)
                    Case 479
                        Call clsA99V001.DIG_BOSTON(Format(gfTratarNulo(pRstApol("Num_Cta_Bas_Cor"), tpoNumerico), "0000000"), lDigito, lAuxiliar)
                    Case 356
                        Call clsA99V001.DIG_REAL(Format(.gAgenciaDig, "0000"), Format(gfTratarNulo(pRstApol("Num_Cta_Bas_Cor"), tpoNumerico), "0000000"), lDigito, lAuxiliar)
                End Select
                Set clsA99V001 = Nothing
                .Num_Cta_Bas_Cor = gfTratarNulo(pRstApol("Num_Cta_Bas_Cor"), tpoNumerico) & lDigito
            Else
                .Num_Cta_Bas_Cor = gfTratarNulo(pRstApol("Num_Cta_Bas_Cor"), tpoNumerico)
            End If
            
            If .Cod_Bco_Cobr = 409 Then
                .Cod_Bco_Cobr = 341
                .Cod_Agen_Bas_Cor = 0
                .Num_Cta_Bas_Cor = 0
            End If
            
            'FICHA 900579 - CARREGAR CAMPOS DO SISCAD
            '------------------------------------------------------------------------------------------------
            If gTipUsuario = e_Tipo_Usu_Corretor Or gTipUsuario = e_Tipo_Usu_CorretorYasuda Then
                On Error Resume Next
                .Tip_Pessoa = gfTratarNulo(pRstApol("Tip_Pessoa"), tpoNumerico)
                .Cod_Ramo_Ativ = gfTratarNulo(pRstApol("Cod_Ramo_Ativ"), tpoNumerico)
                .Cod_Doc = gfTratarNulo(pRstApol("Cod_Doc"), tpoNumerico)
                .Num_Doc = gfTratarNulo(pRstApol("Num_Doc"), TpoString)
                .Org_Exp = gfTratarNulo(pRstApol("Org_Exp"), TpoString)
                .Dat_Emis_Doc = gfTratarNulo(pRstApol("Dat_Emis_Doc"), tpoNumerico)
                'Ficha 900417
                If gCodProduto >= 968 Then
                    .Nom_Empresa = gfTratarNulo(pRstApol("Nom_Profissao"), TpoString)
                Else
                    .Nom_Empresa = gfTratarNulo(pRstApol("Nom_Empresa"), TpoString)
                End If
                .Sexo = gfTratarNulo(pRstApol("Sexo"), TpoString)
                Err.Clear
                On Error GoTo ControleErro
            End If
            '------------------------------------------------------------------------------------------------
            
'Ficha 1000496
             'Correção de Produção - Unidade
             If gBaseLocal = False Then
                 lNomeTab = "GED.DBO.TAB_CORRECAO_PROD"
                 lSql = " Select top 1 cod_unid_acert from " & lNomeTab & " Where Num_Apol = " & pRstApol("Num_Apol")
                 lSql = lSql & " and cod_unid_acert > 0 and niv_corr = 1  "
                 lSql = lSql & " AND DAT_TER  = 99999999"
            
                 If Not gfObtRegistro(pBdP00APOL, lSql, rsAPOL, pMensagem) Then
                     GoTo SAIDA
                 End If
             
                 If Not rsAPOL.EOF Then
                     .Cod_Unid_Prod = gfTratarNulo(rsAPOL("cod_unid_acert"), tpoNumerico)
                 End If
                 
             'Correção de Produção - Produtor
                 lNomeTab = "GED.DBO.TAB_CORRECAO_PROD"
                 lSql = " Select top 1 cod_prod_acert from " & lNomeTab & " Where Num_Apol = " & pRstApol("Num_Apol")
                 lSql = lSql & " and cod_prod_acert > 0 and niv_corr = 3  "
                 lSql = lSql & " AND DAT_TER  = 99999999"
            
                 If Not gfObtRegistro(pBdP00APOL, lSql, rsAPOL, pMensagem) Then
                     GoTo SAIDA
                 End If
             
                 If Not rsAPOL.EOF Then
                     .Cod_Prod = gfTratarNulo(rsAPOL("cod_prod_acert"), tpoNumerico)
                 End If
             End If
             
            If pNumCotac_LMI <> "" Then '1100177 - Renovação Apolice LMI
                .LMI_Unico = 1
                .Num_Apol_An = pNumApolice
            End If
            .Qtd_Parc = 10
        End With
        
        If Not lobjPedido.mfIncluir(lobjstcPedido, lPropTabela, "") Then
            pMensagem = lobjPedido.UltimoErro
            GoTo SAIDA
        End If
        
        If gBaseLocal = True Then
            lNomeTab = "TAB_APOL"
        Else
            If pNumCotac_LMI <> "" Then '1100177 - Renovação Apolice LMI
                lNomeTab = "RAMOSDIVERSOS.DBO.TAB_PED"
            Else
                lNomeTab = "GED.DBO.TAB_PED"
            End If
        End If
        
        'Corretores
        lSql = " Select * from " & lNomeTab & "_CORR Where "
        If pNumCotac_LMI = "" Then '1100177 - Renovação Apolice LMI
            lSql = lSql & " Num_Apol = " & pRstApol("Num_Apol")
            lSql = lSql & " and Num_Endo = " & pRstApol("Num_Endo")
        Else
            lSql = lSql & " NOSSO_NUMERO = " & mfFormatarCamposSQL(pNumCotac_LMI)
        End If
        
        If Not gfObtRegistro(pBdP00APOL, lSql, rsAPOL, pMensagem) Then
            GoTo SAIDA
        End If
        
        Set lobjCorr = New clsA46V705A
        Set lobjCorr.ConexaoBD = lobjPedido.ConexaoBD
        Do While Not rsAPOL.EOF
            Set lObjstcCorr = New stcA46V705B
            With lObjstcCorr
                .Cod_Ramo = gfTratarNulo(rsAPOL("Cod_Ramo"), tpoNumerico)
                .Nosso_Numero = pNosNumero
                .Num_Ped = pNumPedido
                If gNum_Pi = 0 Then
                    .Num_PI = Right(Year(Date), 2) & Format(pNumPedido, "00000000")
                Else
                    .Num_PI = gNum_Pi
                End If
                .Cod_Corr = gfTratarNulo(rsAPOL("Cod_Corr"), tpoNumerico)
                gCod_Corretor = .Cod_Corr '1001135 - Corretor Volks - Franquia Especial Concessionaria
                .Num_Ordem_Corr = gfTratarNulo(rsAPOL("Num_Ordem_Corr"), tpoNumerico)
             '   .Per_Corr = gfTratarNulo(rsApol("Per_Corr"), TpoValor)
                .Per_Corr = 0
                
                If .Num_Ordem_Corr = 1 Then
                'Posiciona o Produtor vigente
                   lobjstcPedido.Cod_Prod = gfRetornaCodProdutor(.Cod_Corr, lobjstcPedido.Cod_Unid_Prod)
                   Call lobjPedido.mfAlterar(lobjstcPedido, lPropTabela, "", , lobjstcPedido.Cod_User_Liber)
                End If
            End With

            Call lobjCorr.mfIncluir(lObjstcCorr, lPropTabela, "")
            rsAPOL.MoveNext
        Loop
        
        'Cosseguro
        lSql = " Select * from " & lNomeTab & "_COSS Where "
        If pNumCotac_LMI = "" Then '1100177 - Renovação Apolice LMI
            lSql = lSql & " Num_Apol = " & pRstApol("Num_Apol")
            lSql = lSql & " and Num_Endo = " & pRstApol("Num_Endo")
        Else
            lSql = lSql & " NOSSO_NUMERO = " & mfFormatarCamposSQL(pNumCotac_LMI)
        End If
        
        If Not gfObtRegistro(pBdP00APOL, lSql, rsAPOL, pMensagem) Then
            GoTo SAIDA
        End If
        
        Set lobjCoss = New clsA46V706A
        Set lobjCoss.ConexaoBD = lobjPedido.ConexaoBD
        Do While Not rsAPOL.EOF
            Set lobjstcCoss = New stcA46V706B
            With lobjstcCoss
                .Cod_Ramo = gfTratarNulo(rsAPOL("Cod_Ramo"), tpoNumerico)
                .Nosso_Numero = pNosNumero
                .Num_Ped = pNumPedido
                If gNum_Pi = 0 Then
                    .Num_PI = Right(Year(Date), 2) & Format(pNumPedido, "00000000")
                Else
                    .Num_PI = gNum_Pi
                End If
                .Cod_Coss = gfTratarNulo(rsAPOL("Cod_Coss"), tpoNumerico)
                .Per_Coss = gfTratarNulo(rsAPOL("Per_Coss"), tpoNumerico)
            End With
            Call lobjCoss.mfIncluir(lobjstcCoss, lPropTabela, "")
            rsAPOL.MoveNext
        Loop
        
        'Endereço
        lSql = " Select * from " & lNomeTab & "_ENDER Where "
        If pNumCotac_LMI = "" Then '1100177 - Renovação Apolice LMI
            lSql = lSql & " Num_Apol = " & pRstApol("Num_Apol")
            lSql = lSql & " and Num_Endo = " & pRstApol("Num_Endo")
        Else
            lSql = lSql & " NOSSO_NUMERO = " & mfFormatarCamposSQL(pNumCotac_LMI)
        End If
        If Not gfObtRegistro(pBdP00APOL, lSql, rsAPOL, pMensagem) Then
            GoTo SAIDA
        End If
        
        Set lObjEnd = New clsA46V707A
        Set lObjEnd.ConexaoBD = lobjPedido.ConexaoBD
        If Not rsAPOL.EOF Then
            Set lObjstcEnd = New stcA46V707B
            With lObjstcEnd
                .Nosso_Numero = pNosNumero
                .Cod_Ramo = lobjstcPedido.Cod_Ramo
                .Num_Ped = lobjstcPedido.Num_Ped
                If gNum_Pi = 0 Then
                    .Num_PI = Right(Year(Date), 2) & Format(pNumPedido, "00000000")
                Else
                    .Num_PI = gNum_Pi
                End If
                .Cod_Tip_Logr = gfTratarNulo(rsAPOL("Cod_Tip_Logr"), TpoString)
                .Nom_Logr = gfTratarNulo(rsAPOL("Nom_Logr"), TpoString)
                .Dsc_Compl_Logr = gfTratarNulo(rsAPOL("Dsc_Compl_Logr"), TpoString)
                .Nom_Bairro = gfTratarNulo(rsAPOL("Nom_Bairro"), TpoString)
                .Nom_Cid = gfTratarNulo(rsAPOL("Nom_Cid"), TpoString)
                .Sig_UF = gfTratarNulo(rsAPOL("Sig_UF"), TpoString)
                .Num_CEP = gfTratarNulo(rsAPOL("Num_Cep"), tpoNumerico)
                .Num_Fone = gfTratarNulo(rsAPOL("Num_Fone"), tpoNumerico)
        'Ficha 1000307
                '.Cod_Reg_Cons = gfTratarNulo(rsAPOL("Cod_Reg_Cons"), tpoNumerico)
                '.Num_Reg_Cons = gfTratarNulo(rsAPOL("Num_Reg_Cons"), tpoString)
            End With
            If Not lObjEnd.mfIncluir(lObjstcEnd, lPropTabela, "") Then
                pMensagem = lObjEnd.UltimoErro
                GoTo SAIDA
            End If
        Else
            Set lObjstcEnd = New stcA46V707B
            With lObjstcEnd
                .Nosso_Numero = pNosNumero
            End With
            If Not lObjEnd.mfIncluir(lObjstcEnd, lPropTabela, "") Then
                pMensagem = lObjEnd.UltimoErro
                GoTo SAIDA
            End If
        End If
        
        'Juros
        'Set lobjJuros = New clsA46V711A
        'Set lobjJuros.ConexaoBD = lobjPedido.ConexaoBD
        'If Not lobjJuros.mfCopiarPadraoParaNovo(pNosNumero) Then
        '    pMensagem = lobjJuros.UltimoErro
        '    GoTo ControleErro
        'End If
        
        'Controle de emissão
        Set lobjCtrlEmis = New clsA46V712A
        Set lobjCtrlEmis.ConexaoBD = lobjPedido.ConexaoBD
        Set lobjstcCtrlEmis = New stcA46V712B
        lobjstcCtrlEmis.Nosso_Numero = pNosNumero
        
        If gCod_Ramo = e_CodRamo_Riscos_diversos Then
            lobjstcCtrlEmis.Num_PI = gNum_Pi
            lobjstcCtrlEmis.Cod_Ramo = e_CodRamo_Riscos_diversos
            lobjstcCtrlEmis.Num_Apol_Ant = lobjstcPedido.Num_Apol_An
        End If
        
        If Not lobjCtrlEmis.mfIncluir(lobjstcCtrlEmis, lPropTabela, "") Then
            pMensagem = lobjCtrlEmis.UltimoErro
            GoTo SAIDA
        End If
        
        'Especificação
        'Set lobjEspec = New clsA46V721A
        'Set lobjEspec.ConexaoBD = lobjPedido.ConexaoBD
        'Set lobjstcEspec = New stcA46V721B
        'lobjstcEspec.Nosso_Numero = pNosNumero
        'If Not lobjEspec.mfIncluir(lobjstcEspec) Then
        '    pMensagem = lobjEspec.UltimoErro
        '    GoTo SAIDA
        'End If
        
        
    Else
        pMensagem = "FGE0089 - Dados de comum não localizados."
    End If
    mfGraRenComum = True
    
SAIDA:
    GoSub FechaObj
    Exit Function
    
ControleErro:
    Call gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gpGraLog(0, "gfGraRenComum ERRO - " & Err & " " & Err.Description & " " & pMensagem)
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
                    

Private Function mfGraRenItem(prstItem As ADODB.Recordset, pNosNumero As String, _
                    pBdP00APOL As ADODB.Connection, pMensagem As String, _
                    pIndSinistro As e_SimNao, pNumPed As Long, pCod_Dizer As Integer, pNumCotac_LMI As String, ByVal pNumApolice As String) As Boolean
    'Função: Grava os registros de comum com os dados da apólice (Renovação Yasuda).
    'Parâmetros de entrada :prstItem.........Registro de item(apólice)
    '                       pNosNumero.......Nosso número.
    '                       pBdP00APOL.......Arquivo de conexão
    'Parâmetro de saída:    pMensagem........Mensagem de erro.
    
    Dim lSql            As String                   'Consulta SQL.
    Dim rsAPOL          As ADODB.Recordset          'Registro.

    '[FRH]->
    'Conferir coberturas na tabela POOMULTI:TAB_COB_RELACAO
    Dim lSql_buscaCob As String
    Dim rs_buscaCob As ADODB.Recordset
    Dim bdP00MULTI As ADODB.Connection
    '<-[FRH]
    
    Dim lObjPedLoc      As clsA46V708A              'Objeto da classe clsA46V708 (TAB_PED_LOC).
    Dim lobjstcPedLoc   As stcA46V708B              'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC).
    Dim lObjClau        As clsA46V703A              'Objeto da classe clsA46V703A (TAB_PED_CLAU).
    Dim lobjstcClau     As stcA46V703B              'Objeto da classe de estrutura stcA46V703B (TAB_PED_CLAU).
    Dim lobjCobert      As clsA46V704A              'Objeto da classe clsA46V704A (TAB_PED_COBERT).
    Dim lobjstcCobert   As stcA46V704B              'Objeto da classe de estrutura stcA46V704B (TAB_PED_COBERT).
    Dim lobjInsp        As clsA46V717A              'Objeto da classe clsA46V717A (TAB_PED_INSP).
    Dim lObjstcInsp     As stcA46V717B              'Objeto da classe de estrutura stcA46V717B (TAB_PED_INSP).
    Dim lNumero         As String
    Dim lComplemento    As String
    Dim lPos            As Long
    Dim lNomeTab        As String
    Dim lValorRisco     As Double
    Dim lObjPedFator    As clsA46V720A
    Dim lobjstcPedFator As stcA46V720B
    Dim lPropTabela     As String
    
    On Error GoTo ControleErro
    Call gfObtSYASMULTINI(lPropTabela)
    
    If gBaseLocal = True Then
        lNomeTab = "TAB_APOL"
    Else
        If pNumCotac_LMI <> "" Then '1100177 - Renovação Apolice LMI
            lNomeTab = "RAMOSDIVERSOS.DBO.TAB_PED"  '1100177 - Renovação Apolice LMI
        Else
            lNomeTab = "GED.DBO.TAB_PED"
        End If
    End If
    
    Set lObjPedLoc = New clsA46V708A
    Set lobjstcPedLoc = New stcA46V708B
    With lobjstcPedLoc
        .Nosso_Numero = pNosNumero
        .Num_Ped = pNumPed
        If gNum_Pi = 0 Then
            .Num_PI = Right(Year(Date), 2) & Format(pNumPed, "00000000")
        Else
            .Num_PI = gNum_Pi
        End If
        .Cod_Ramo = gfTratarNulo(prstItem("Cod_Ramo"), tpoNumerico)
        .Num_Item = gfTratarNulo(prstItem("Num_item"), tpoNumerico)
        .Cod_Iden_Tarif = gCodProduto
        .Cod_Tip_Logr_Risco = gfTratarNulo(prstItem("Cod_Tip_Logr_Risco"), TpoString)
        .Nom_Logr_Risco = gfTratarNulo(prstItem("Nom_Logr_Risco"), TpoString)
        lComplemento = gfTratarNulo(prstItem("Dsc_Compl_Risco"), TpoString)
        
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
        .Dsc_Compl_Risco = Left(lComplemento, 20)
        .Nom_Bairro_Risco = gfTratarNulo(prstItem("Nom_Bairro_Risco"), TpoString)
        .Nom_Cid_Risco = gfTratarNulo(prstItem("Nom_Cid_Risco"), TpoString)
        .Sig_UF_Risco = gfTratarNulo(prstItem("Sig_UF_Risco"), TpoString)
        .Num_Cep_Risco = gfTratarNulo(prstItem("Num_CEP_Risco"), tpoNumerico)
        .Per_Desc_Apar = gfTratarNulo(prstItem("Per_Desc_Apar"), TpoValor)
        
        If gCod_Plano >= 2021 And gCod_Plano <= 2031 Then
            .Per_Desc_Progres = 0
        ElseIf gCod_Plano >= 2061 And gCod_Plano <= 2100 Then
            .Per_Desc_Progres = 0
    'Ficha 901673
        ElseIf gCod_Plano >= 2111 And gCod_Plano <= 2120 Then
            .Per_Desc_Progres = 0
        Else
            .Per_Desc_Progres = gfTratarNulo(prstItem("Per_Desc_Progres"), TpoValor)
        End If
        .Per_IS_Vr = gfTratarNulo(prstItem("Per_IS_Vr"), TpoValor)
        .Val_Coef_IS_Vr = gfTratarNulo(prstItem("Val_Coef_IS_Vr"), TpoValor)
        .Val_Vr = gfTratarNulo(prstItem("Val_Vr"), TpoValor)
        lValorRisco = .Val_Vr
        .Tab_Clasf = gfTratarNulo(prstItem("Tab_Clasf"), tpoNumerico)
        .Cod_Clasf = gfTratarNulo(prstItem("Cod_Clasf"), TpoString)  '1000871 - Inclusão de Orgão Categoria Conselho
        .Cod_Compl_Clasf = gfTratarNulo(prstItem("Cod_Compl_Clasf"), TpoString) '1000871 - Inclusão de Orgão Categoria Conselho
        .Dsc_Clasf = gfTratarNulo(prstItem("Dsc_Clasf"), TpoString)
        .Dsc_Clasf = ""
        .Dsc_Clasf = Left(.gNom_Class, 60)
        .Cod_Constr = gfTratarNulo(prstItem("Cod_Constr"), tpoNumerico)
        .Cod_User_Liber = gNomUsuario6
        .Dat_Liber = Left(gfForData(3), 8)
        .Hor_Liber = Mid(gfForData(3), 9, 4)
        .Cod_Sit_item = e_SitItem_NaoCalc
        .Per_Bonus = gfTratarNulo(prstItem("Per_Bonus"), tpoNumerico)
        .Num_Apol_Anterior = prstItem("Num_APol")
        .Num_Item_Anterior = prstItem("Num_Item")
        
        If pNumCotac_LMI <> "" Then '1100177 - Renovação Apolice LMI
            .Num_Apol_Anterior = pNumApolice
        End If
        
        If lobjstcPedLoc.mBonus_Anterior = 0 Then
            .Per_Bonus = lObjPedLoc.mfRetornaBonusApolice(lobjstcPedLoc)
        Else
            .Per_Bonus = lobjstcPedLoc.mBonus_Anterior
        End If
        
        If pIndSinistro = e_SimNao_Sim Then
            .Per_Bonus = .Per_Bonus - 5
        Else
            .Per_Bonus = .Per_Bonus + 5
        End If
        If .Per_Bonus < 0 Then .Per_Bonus = 0
                
        If .Cod_Ramo = e_CodRamo_Empresarial Then
            If gCodProduto >= 942 Then
                    'Soma mais 5 da fidelidade.
                .Per_Bonus = .Per_Bonus + 5
            End If
            If .Per_Bonus > 30 Then .Per_Bonus = 30
        ElseIf .Cod_Ramo = e_CodRamo_Residencial Then
            If gCodProduto >= 946 Then
                    'Soma mais 5 da fidelidade.
                .Per_Bonus = .Per_Bonus + 5
                If .Per_Bonus > 30 Then .Per_Bonus = 30
            Else
                If .Per_Bonus > 25 Then .Per_Bonus = 25
            End If
        Else
            If .Per_Bonus > 25 Then .Per_Bonus = 25
        End If
        
        If gCod_Plano >= 2001 And gCod_Plano <= 2031 Then
            .Per_Bonus = 0
        End If
        
    'Ficha 901673
        If (gCod_Plano >= 2051 And gCod_Plano <= 2100) Or (gCod_Plano >= 2111 And gCod_Plano <= 2120) Then
            .Per_Bonus = 0
        End If

'Ficha 900550
        gDesc_Bonus = .Per_Bonus
        
        .Qtd_Empreg = gfTratarNulo(prstItem("Qtd_Empreg"), tpoNumerico)
        .Cod_Iden_Segur_01_V = gfTratarNulo(prstItem("Cod_Iden_Segur_01"), TpoString)
        .Cod_Iden_Segur_02_A = gfTratarNulo(prstItem("Cod_Iden_Segur_02"), TpoString)
        .Cod_Iden_Segur_03_P = gfTratarNulo(prstItem("Cod_Iden_Segur_03"), TpoString)
        .Cod_Iden_Segur_04_I = gfTratarNulo(prstItem("Cod_Iden_Segur_04"), TpoString)                    'Inexistência de terreno baldio,  imóvel desocupado   ou em construção na vizinhança do risco.: S ou N
        .Cod_Iden_Segur_05_G = gfTratarNulo(prstItem("Cod_Iden_Segur_05"), TpoString)                    'Grades metálicas de proteção em todas as janelas e portas de vidros que dão acesso ao interior do imóvel desde que também possuam alarme.: S ou N
        .Per_Tot_Segur = .gPer_Tot_Segur
        .Cod_Iden_Protec_01_E = gfTratarNulo(prstItem("Cod_Iden_Protec_01"), TpoString)
        .Cod_Iden_Protec_02_H = gfTratarNulo(prstItem("Cod_Iden_Protec_02"), TpoString)
        .Cod_Iden_Protec_03_S = gfTratarNulo(prstItem("Cod_Iden_Protec_03"), TpoString)
        .Cod_Iden_Protec_04_D = gfTratarNulo(prstItem("Cod_Iden_Protec_04"), TpoString)                   'Indicador de Detecção e alarmes: S ou N.
        .Per_Tot_Protec = .gPer_Tot_Protec
        If gCod_Plano > 1010 Then
            .Cod_Produto = gfTratarNulo(prstItem("Cod_Produto"), tpoNumerico)
        Else
            .Cod_Produto = 1
        End If
            
        .Cod_Regiao = gfTratarNulo(prstItem("Cod_Regiao"), tpoNumerico)
        If pCod_Dizer = 0 Then
            .Tip_Emissao = e_TipEmissao_RenYasuda
        Else
            .Tip_Emissao = e_TipEmissao_Endosso
        End If
        .Ind_Sinistro = pIndSinistro
        
''Ficha 1000305
'        .Tip_Cobert = gfTratarNulo(prstItem("Tip_Cobert"), tpoNumerico)
'        gTipCobertura = .Tip_Cobert
        
    End With
    
    If Not lObjPedLoc.mfIncluir(lobjstcPedLoc, lPropTabela, "") Then
        pMensagem = lObjPedLoc.UltimoErro
        Exit Function
    End If
    
    'Clausulas
    lSql = " Select * from " & lNomeTab & "_CLAU "
    If gBaseLocal = False And pBdP00APOL Like "*Provider=MSDASQL.*" Then
        lSql = lSql & " inner join " & IIf(pNumCotac_LMI = "", "GED", "RAMOSDIVERSOS") & ".DBO.tab_ctrl_emis  as ctrl on " & lNomeTab & "_CLAU.num_apol = ctrl.num_apol"
        lSql = lSql & " and ctrl.num_endo = " & lNomeTab & "_CLAU.num_endo"
        lSql = lSql & " and ctrl.num_PI = " & lNomeTab & "_CLAU.num_PI"
    End If
    If pNumCotac_LMI <> "" Then '1100177 - Renovação Apolice LMI
        lSql = lSql & " Where " & lNomeTab & "_CLAU.NOSSO_NUMERO = " & mfFormatarCamposSQL(pNumCotac_LMI)
        lSql = lSql & " and ctrl.NOSSO_NUMERO = " & lNomeTab & "_CLAU.NOSSO_NUMERO"
    Else
        lSql = lSql & " Where " & lNomeTab & "_CLAU.Num_Apol = " & prstItem("Num_Apol")
        lSql = lSql & " and " & lNomeTab & "_CLAU.num_Endo = " & prstItem("NUm_Endo2")
    End If
    lSql = lSql & " and " & lNomeTab & "_CLAU.num_item = " & prstItem("Num_Item")
    If Not gfObtRegistro(pBdP00APOL, lSql, rsAPOL, pMensagem) Then
        GoTo ControleErro
    End If
    
    Set lObjClau = New clsA46V703A
    Set lObjClau.ConexaoBD = lObjPedLoc.ConexaoBD
    Do While Not rsAPOL.EOF
        Set lobjstcClau = New stcA46V703B
        With lobjstcClau
            .Cod_Ramo = gfTratarNulo(rsAPOL("Cod_Ramo"), tpoNumerico)
            .Nosso_Numero = pNosNumero
            .Num_Ped = pNumPed
            If gNum_Pi = 0 Then
                .Num_PI = Right(Year(Date), 2) & Format(pNumPed, "00000000")
            Else
                .Num_PI = gNum_Pi
            End If
            .Cod_Clau = gfTratarNulo(rsAPOL("Cod_Clau"), tpoNumerico)
            .Num_Item = lobjstcPedLoc.Num_Item
        End With
        If Not lObjClau.mfIncluir(lobjstcClau, lPropTabela, "") Then
            pMensagem = lObjClau.UltimoErro
            Exit Function
        End If
        rsAPOL.MoveNext
    Loop
    
    'COberturas
    lSql = " Select " & lNomeTab & "_COBERT.* from " & lNomeTab & "_COBERT "
    If gBaseLocal = False And pBdP00APOL Like "*Provider=MSDASQL.*" Then
        lSql = lSql & " inner join " & IIf(pNumCotac_LMI = "", "GED", "RAMOSDIVERSOS") & ".DBO.tab_ctrl_emis  as ctrl on " & lNomeTab & "_COBERT.num_apol = ctrl.num_apol"
        lSql = lSql & " and ctrl.num_endo = " & lNomeTab & "_COBERT.num_endo"
        lSql = lSql & " and ctrl.num_PI = " & lNomeTab & "_COBERT.num_PI"
    End If
    If pNumCotac_LMI <> "" Then '1100177 - Renovação Apolice LMI
        lSql = lSql & " Where " & lNomeTab & "_COBERT.NOSSO_NUMERO = " & mfFormatarCamposSQL(pNumCotac_LMI)
        lSql = lSql & " and ctrl.NOSSO_NUMERO = " & lNomeTab & "_COBERT.NOSSO_NUMERO"
    Else
        lSql = lSql & " Where " & lNomeTab & "_COBERT.Num_Apol = " & prstItem("Num_Apol")
        lSql = lSql & " and " & lNomeTab & "_COBERT.num_Endo = " & prstItem("NUm_Endo2")
    End If
    lSql = lSql & " and " & lNomeTab & "_COBERT.num_item = " & prstItem("Num_Item")
    If prstItem("Cod_Ramo") = e_CodRamo_Riscos_diversos Then
        lSql = lSql & " order by GED.dbo.TAB_PED_COBERT.tab_cobert, GED.dbo.TAB_PED_COBERT.cod_cobert "
    End If
    If Not gfObtRegistro(pBdP00APOL, lSql, rsAPOL, pMensagem) Then
        GoTo ControleErro
    End If
    
    Set lobjCobert = New clsA46V704A
   'Set lObjCobert.ConexaoBD = lObjPedLoc.ConexaoBD
    '[FRH]->
    If Not gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, bdP00MULTI, pMensagem) Then
    GoTo ControleErro
    End If
    '<-[FRH]
    Do While Not rsAPOL.EOF
        Set lobjstcCobert = New stcA46V704B
        With lobjstcCobert
            .Cod_Ramo = gfTratarNulo(rsAPOL("Cod_Ramo"), tpoNumerico)
            .Nosso_Numero = pNosNumero
            .Num_Ped = pNumPed
            If gNum_Pi = 0 Then
                .Num_PI = Right(Year(Date), 2) & Format(pNumPed, "00000000")
            Else
                .Num_PI = gNum_Pi
            End If
            .Num_Item = lobjstcPedLoc.Num_Item
                        
            '[FRH]->
            'Rotina para conferir alterações de códigos de cobertura,
            'na *POOMULTI.TAB_COB_RELACAO* existe um de-para de cobs.anteriores e atuais.
            
            '[Valor original Cod_Cobert da TAB_APOL]
            .Cod_Cobert = gfTratarNulo(rsAPOL("cod_cobert"), tpoNumerico)
            
            lSql_buscaCob = "SELECT COD_COB_ATUAL FROM TAB_COB_RELACAO WHERE COD_COB_ANTERIOR=" & gfTratarNulo(rsAPOL("Cod_Cobert"), tpoNumerico)
            If gfObtRegistro(bdP00MULTI, lSql_buscaCob, rs_buscaCob, pMensagem) Then
                If Not (rs_buscaCob.BOF And rs_buscaCob.EOF) Then
                    .Cod_Cobert = gfTratarNulo(rs_buscaCob("cod_cob_atual"), tpoNumerico)
                End If
            End If
            '<-[FRH]
            
            If .Cod_Ramo = e_CodRamo_Riscos_diversos Then
                .Tab_Cobert = gfTratarNulo(rsAPOL("Tab_Cobert"), tpoNumerico)
                gCod_Plano = .Tab_Cobert
            End If
            If gCod_Plano > 1010 Then
                .Tab_Cobert = gCod_Plano
            Else
                .Tab_Cobert = 1
            End If
            
            'Fichas 901282 e 901346
            Select Case gCod_Plano
'1000744 - Item 4 - Plano Mercedes
                Case 2021 To 2031
                    If gCodProduto < 992 Then
                        If lobjstcPedLoc.Cod_Clasf = "001" And lobjstcPedLoc.Cod_Compl_Clasf = "11" Then .Per_Desc_Coml = 33
                        If lobjstcPedLoc.Cod_Clasf = "001" And lobjstcPedLoc.Cod_Compl_Clasf = "21" Then .Per_Desc_Coml = 20
                    End If
                Case 2051 To 2060
                    If lobjstcPedLoc.Cod_Clasf = "001" And lobjstcPedLoc.Cod_Compl_Clasf = "11" Then .Per_Desc_Coml = 20
                    If lobjstcPedLoc.Cod_Clasf = "001" And lobjstcPedLoc.Cod_Compl_Clasf = "21" Then .Per_Desc_Coml = 5
                Case 2061 To 2100
                    If lobjstcPedLoc.Cod_Clasf = "001" And lobjstcPedLoc.Cod_Compl_Clasf = "11" Then .Per_Desc_Coml = 20
                    If lobjstcPedLoc.Cod_Clasf = "001" And lobjstcPedLoc.Cod_Compl_Clasf = "21" Then .Per_Desc_Coml = 5
                'Ficha 901673
                Case 2111 To 2120
                    If lobjstcPedLoc.Cod_Clasf = "001" And lobjstcPedLoc.Cod_Compl_Clasf = "11" Then .Per_Desc_Coml = 20
                    If lobjstcPedLoc.Cod_Clasf = "001" And lobjstcPedLoc.Cod_Compl_Clasf = "21" Then .Per_Desc_Coml = 5
                Case Else
                    .Per_Desc_Coml = gfTratarNulo(rsAPOL("Per_Desc_Coml"), tpoNumerico)
                    '1200539 - Tarifa 1038 e 1039 Junho 2012 - Item 1
                    Dim lCoef_Desc_Coml As Double
                    If lobjstcPedLoc.Cod_Iden_Tarif >= 1038 Then
                        If mfAjustesPontuais(e_Tip_Valor_DescComl, TpoValor, lCoef_Desc_Coml, pMensagem, lobjstcPedLoc) = False Then
                           GoTo ControleErro
                        End If
                        If .Per_Desc_Coml < lCoef_Desc_Coml Then
                            .Per_Desc_Coml = Format(lCoef_Desc_Coml, "standard")
                        End If
                    End If
            End Select
         
'Ficha 900550
'            .Per_Desc_Exp = gfTratarNulo(rsAPOL("Per_Desc_Exp"), tpoNumerico)
                    
            .Per_Desc_Exp = gDesc_Bonus
                       
            .Val_IS = gfTratarNulo(rsAPOL("Val_IS"), tpoNumerico)
            .Val_Risco_Cob = .Val_IS
            
            If pCod_Dizer <> 0 Then
                If gCod_Plano > 1010 And .Cod_Ramo = e_CodRamo_Residencial Then
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
            Call mpBuscaLimFranquia(lobjstcPedLoc, lobjstcCobert)
            '11/01/2010 - Ficha: 1000003 - franquia e valor de franquia - fim
            
        End With

        If Not lobjCobert.mfIncluir(lobjstcCobert, lPropTabela, "") Then
            If UCase(lobjCobert.UltimoErro) Like "*DUPLICATE*" And Not rs_buscaCob.EOF And Not rs_buscaCob.BOF Then
                GoTo Proximo
            ElseIf UCase(lobjCobert.UltimoErro) Like "*DUPLICATE*" And lobjstcCobert.Cod_Ramo = e_CodRamo_Riscos_diversos Then
                GoTo Proximo
            Else
                pMensagem = lobjCobert.UltimoErro
                Exit Function
            End If
        End If
        If lobjstcCobert.Cod_Cobert = 1 And lValorRisco = 0 Then
            lobjstcPedLoc.Val_Vr = lobjstcCobert.Val_IS
            lobjstcPedLoc.Per_IS_Vr = 100
            If Not lObjPedLoc.mfAlterarValorRisco(lobjstcPedLoc, lPropTabela, "") Then
                pMensagem = lObjPedLoc.UltimoErro
                Exit Function
            End If
        End If
        
Proximo:
        rsAPOL.MoveNext
    Loop
        
    '[FRH]->
    Call gpFecha2(bdP00MULTI)
    '<-[FRH]
    
    Set lobjInsp = New clsA46V717A
    Set lObjstcInsp = New stcA46V717B
    lObjstcInsp.Nosso_Numero = pNosNumero
    lObjstcInsp.Num_Item = lobjstcPedLoc.Num_Item
    lObjstcInsp.Num_Logr_Risco = Val(lNumero)
    'correção
    lObjstcInsp.Cod_Ramo = lobjstcPedLoc.Cod_Ramo
    lObjstcInsp.Num_Ped = lobjstcPedLoc.Num_Ped
    '
    If Not lobjInsp.mfIncluir(lObjstcInsp, lPropTabela, "") Then
        pMensagem = lobjInsp.UltimoErro
        Exit Function
    End If
    
        
    'Ficha 1000409 - Inclusão de Fatores de Proteção em Renovações Yasuda
    If gBaseLocal = True Then
        lNomeTab = "TAB_APOL"
    Else
        lNomeTab = "RamosDiversos.dbo.Tab_Ped"
    End If
    lSql = " Select * from " & lNomeTab & "_FAT_RISCO "
    If gBaseLocal = False And pBdP00APOL Like "*Provider=MSDASQL.*" Then
        lSql = lSql & " inner join " & IIf(pNumCotac_LMI = "", "GED", "RAMOSDIVERSOS") & ".DBO.tab_ctrl_emis  as ctrl on " & lNomeTab & "_FAT_RISCO.num_apol = ctrl.num_apol"
        lSql = lSql & " and ctrl.num_endo = " & lNomeTab & "_FAT_RISCO.num_endo"
        lSql = lSql & " and ctrl.num_PI = " & lNomeTab & "_FAT_RISCO.num_PI"
    End If
    If pNumCotac_LMI <> "" Then '1100177 - Renovação Apolice LMI
        lSql = lSql & " Where " & lNomeTab & "_FAT_RISCO.NOSSO_NUMERO = " & mfFormatarCamposSQL(pNumCotac_LMI)
        lSql = lSql & " and ctrl.NOSSO_NUMERO = " & lNomeTab & "_FAT_RISCO.NOSSO_NUMERO"
    Else
        lSql = lSql & " Where " & lNomeTab & "_FAT_RISCO.Num_Apol = " & prstItem("Num_Apol")
        lSql = lSql & " and " & lNomeTab & "_FAT_RISCO.num_Endo = " & prstItem("NUm_Endo2")
    End If
    lSql = lSql & " and " & lNomeTab & "_FAT_RISCO.num_item = " & prstItem("Num_Item")
    On Error Resume Next
    If Not gfObtRegistro(pBdP00APOL, lSql, rsAPOL, pMensagem) Then
        GoTo ControleErro
    End If
    
    Set lObjPedFator = New clsA46V720A
    Set lObjPedFator.ConexaoBD = lObjPedLoc.ConexaoBD
    Do While Not rsAPOL.EOF
        Set lobjstcPedFator = New stcA46V720B
        With lobjstcPedFator
            .Cod_Ramo = gfTratarNulo(rsAPOL("Cod_Ramo"), tpoNumerico)
            .Nosso_Numero = pNosNumero
            .Num_Ped = pNumPed
            If gNum_Pi = 0 Then
                .Num_PI = Right(Year(Date), 2) & Format(pNumPed, "00000000")
            Else
                .Num_PI = gNum_Pi
            End If
            .Num_Item = lobjstcPedLoc.Num_Item
            .Cod_Fator = gfTratarNulo(rsAPOL("Cod_Fator"), tpoNumerico)
            .Dsc_Compl_Fator = gfTratarNulo(rsAPOL("DSC_COMPL_FATOR"), tpoNumerico)
        End With
        If Not lObjPedFator.mfIncluir(lobjstcPedFator, lPropTabela, "") Then
            pMensagem = lObjPedFator.UltimoErro
            Exit Function
        End If
        rsAPOL.MoveNext
    Loop
    On Error GoTo ControleErro
    
'gravar fator de proteção
    If lobjstcPedLoc.Cod_Ramo = e_CodRamo_Residencial And _
      ((gCodProduto >= 1022 And gCodProduto <= 1033) Or gCodProduto >= 1038) Then   '1200539 - Tarifa 1038 e 1039 Junho 2012 - Item 8
      
        Set lobjstcPedFator = New stcA46V720B
        Set lObjPedFator = New clsA46V720A
        With lobjstcPedFator
            .Nosso_Numero = lobjstcPedLoc.Nosso_Numero
            .Num_Ped = lobjstcPedLoc.Num_Ped
            .Cod_Ramo = lobjstcPedLoc.Cod_Ramo
            .Num_Item = lobjstcPedLoc.Num_Item
            .Cod_Org = lobjstcPedLoc.Cod_Org
            .Cod_Fator = 7
            .Desc_Fator = "DESCONTO ESPECIAL"
            Set .ObjPai = lobjstcPedLoc
        End With
    
        Call lObjPedFator.mfIncluir(lobjstcPedFator, lPropTabela, "") 'Ignora Erros!!!
        
    End If


    Set lObjClau = Nothing
    Set lobjCobert = Nothing
    Set lObjPedLoc = Nothing
    Set lobjstcClau = Nothing
    Set lobjstcCobert = Nothing
    Set lobjstcPedLoc = Nothing
    
    mfGraRenItem = True
    Exit Function
ControleErro:
    Set lObjClau = Nothing
    Set lobjCobert = Nothing
    Set lObjPedLoc = Nothing
    Set lobjstcClau = Nothing
    Set lobjstcCobert = Nothing
    Set lobjstcPedLoc = Nothing
    Call gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gpGraLog(0, "gfGraRenItem ERRO - " & Err & " " & Err.Description & " " & pMensagem)
End Function
                    



'11/01/2010 - Ficha: 1000003 - franquia e valor de franquia
Private Sub mpBuscaLimFranquia(ByRef pobjstcPedLoc As stcA46V708B, ByRef pobjstcCobert As stcA46V704B)
Dim BDP00Tarifa         As ADODB.Connection
Dim rstTabClas          As ADODB.Recordset
Dim lbdTbm_Fat_Franquia As ADODB.Connection         'Arquivo: P00Multi.
Dim lMensagem           As String                   'Mensagem.
Dim lrsTbm_Fat_Franquia As ADODB.Recordset          'Registro : Tbm_Fator_Franquia.
Dim lCodSequencia       As Integer
Dim lCEP                As String
Dim lTip_Atividade      As Integer          'Utilizado para selecao de Franquia
Dim lSelect             As String

    pobjstcCobert.Per_Franq = 0
    pobjstcCobert.Val_Franq_Min = 0
    pobjstcCobert.Val_Fator_Franq = 0
  
        
    If Trim(pobjstcPedLoc.Cod_Clasf) <> "" And Trim(pobjstcPedLoc.Cod_Clasf) <> "0" Then
        '1. Abre arquivo.
        If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, BDP00Tarifa, lMensagem) = False Then
            Call gpGraLog(1, lMensagem)
            Exit Sub
        End If
    
        '3. Obtém produto sem término de vigência ou com término de vigência superior à data do dia.
        lSelect = " Select distinct dsc_clasf, COD_CLASF, " & _
                "COD_COMPL_CLASF , tab_Clasf, Tip_Atividade  from Tbm_Classificacao " & _
                " Where cod_Ramo  = " & pobjstcPedLoc.Cod_Ramo & _
                " and cod_Clasf = '" & pobjstcPedLoc.Cod_Clasf & "'" & _
                " and Cod_Compl_Clasf = '" & pobjstcPedLoc.Cod_Compl_Clasf & "'" & _
                " and Cod_Tarif = " & pobjstcPedLoc.Cod_Iden_Tarif & _
                " and Cod_Sequencia <= " & gCodSequencia
        lSelect = lSelect & " and  tab_clasf = " & gCod_Plano
        
        If gfObtRegistro(BDP00Tarifa, lSelect, rstTabClas, lMensagem) = False Then
            Call gpFecha2(BDP00Tarifa)
            Call gpGraLog(1, lMensagem)
            Exit Sub
        End If
    
        If Not rstTabClas.EOF Then
            lTip_Atividade = rstTabClas("Tip_Atividade")
        End If
    End If
    
    
    
    
    If gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, lbdTbm_Fat_Franquia, lMensagem) = False Then
        Call gpGraLog(2, lMensagem)
        Exit Sub
    End If
    
    lCodSequencia = gCodSequencia
    
NovaSequencia1:
    
    '1 - Procura Registro de exceção com a chave completa
    'Procura o valor do Val_Franq_Min e Perc_Franq
    lCEP = Val(Trim(Replace(pobjstcPedLoc.Num_Cep_Risco, "-", "")))
    
    lSelect = " Select Val_Franq_Min as Min , Perc_Franquia as Perc, Fator_Ajuste as Fator from Tbm_Fator_Franquia "
    lSelect = lSelect & " Where Cod_Ramo = " & pobjstcPedLoc.Cod_Ramo
    lSelect = lSelect & " and Cod_Modalidade = " & gCod_Plano
    lSelect = lSelect & " and Cod_Tip_Reg = 'E' "
    If gCodProduto >= 940 Then
        lSelect = lSelect & " and ((Cod_Clasf = '" & pobjstcPedLoc.Cod_Clasf & "'"
        lSelect = lSelect & " and Cod_Compl_Clasf = '" & pobjstcPedLoc.Cod_Compl_Clasf & "')"
        lSelect = lSelect & " or (Cod_Clasf = '0'))"
    Else
        lSelect = lSelect & " and Cod_Clasf = '" & pobjstcPedLoc.Cod_Clasf & "'"
        lSelect = lSelect & " and Cod_Compl_Clasf = '" & pobjstcPedLoc.Cod_Compl_Clasf & "'"
    End If
    lSelect = lSelect & " and Cod_Cobert = " & pobjstcCobert.Cod_Cobert
    lSelect = lSelect & " and Cod_Construcao in (0 , " & pobjstcPedLoc.Cod_Constr & ")"
    lSelect = lSelect & " and Cod_Tarif = " & pobjstcPedLoc.Cod_Iden_Tarif
    lSelect = lSelect & " and Cod_Sequencia <= " & lCodSequencia
    
    If gCodProduto >= 940 Then
        lSelect = lSelect & " AND ((CEP_INICIO <= " & lCEP
        lSelect = lSelect & " AND CEP_FIM >= " & lCEP & " ) "
        lSelect = lSelect & " OR (CEP_INICIO =  0 ))"
        lSelect = lSelect & " order by Prioridade desc "
    End If

    If gfObtRegistro(lbdTbm_Fat_Franquia, lSelect, lrsTbm_Fat_Franquia, lMensagem) = False Then
        Call gpFecha2(lbdTbm_Fat_Franquia)
        Call gpGraLog(2, lMensagem)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    '2 - Se não encontrar a chave completa vai procurar o padrão
    If lrsTbm_Fat_Franquia.EOF Then
        If lCodSequencia > 0 Then
            lCodSequencia = lCodSequencia - 1
            lrsTbm_Fat_Franquia.Close
            GoTo NovaSequencia1
        Else
            GoTo ProcuraDefault
        End If
    End If
    
    If lrsTbm_Fat_Franquia(0) = 0 Then
        GoTo ProcuraDefault
    End If
    
    pobjstcCobert.Per_Franq = lrsTbm_Fat_Franquia("Perc")
    pobjstcCobert.Val_Franq_Min = lrsTbm_Fat_Franquia("Min")
    pobjstcCobert.Val_Fator_Franq = lrsTbm_Fat_Franquia("Fator")
        
    GoTo FimBusca

ProcuraDefault:

    lrsTbm_Fat_Franquia.Close
  
    '3 - Procura o registro somente com Ramo, Modalidade, Cobertura e TipoAtividade
    'Procura o valor do Val_Franq_Min, Perc_Franq
    
    lCodSequencia = gCodSequencia
    
NovaSequencia2:
    
    lSelect = " Select Val_Franq_Min as Min , Perc_Franquia as Perc, Fator_Ajuste as Fator  from Tbm_Fator_Franquia "
    lSelect = lSelect & " Where Cod_Ramo = " & pobjstcPedLoc.Cod_Ramo
    lSelect = lSelect & " and Cod_Modalidade = " & gCod_Plano
    lSelect = lSelect & " and Cod_Tip_Reg = 'D' "
    lSelect = lSelect & " and Cod_Cobert = " & pobjstcCobert.Cod_Cobert
    lSelect = lSelect & " and Tip_Atividade in (3," & lTip_Atividade & ")"
    lSelect = lSelect & " and Cod_Tarif = " & pobjstcPedLoc.Cod_Iden_Tarif
    lSelect = lSelect & " and Cod_Sequencia <= " & lCodSequencia

    If gfObtRegistro(lbdTbm_Fat_Franquia, lSelect, lrsTbm_Fat_Franquia, lMensagem) = False Then
        Call gpFecha2(lbdTbm_Fat_Franquia)
        Call gpGraLog(2, lMensagem)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    If lrsTbm_Fat_Franquia.EOF Then
        If lCodSequencia > 0 Then
            lCodSequencia = lCodSequencia - 1
            lrsTbm_Fat_Franquia.Close
            GoTo NovaSequencia2
        Else
            GoTo FimBusca
        End If
    End If
    
    If lrsTbm_Fat_Franquia(0) = 0 Then
        GoTo FimBusca
    End If

    pobjstcCobert.Per_Franq = lrsTbm_Fat_Franquia("Perc")
    pobjstcCobert.Val_Franq_Min = lrsTbm_Fat_Franquia("Min")
    pobjstcCobert.Val_Fator_Franq = lrsTbm_Fat_Franquia("Fator")

FimBusca:
    Call gpFecha2(lbdTbm_Fat_Franquia)
    
    '1001135 - Corretor Volks - Franquia Especial Concessionaria
    If gCodProduto >= 1002 Then
        If (InStr(1, "00008001/00008002/00008003/00008004/00008005/00008006/00008008/00008011/00008014/00008028", Format(gCod_Corretor, "00000000"), vbTextCompare) > 0) And _
           (Val(pobjstcPedLoc.Cod_Clasf) = 22 And Val(pobjstcPedLoc.Cod_Compl_Clasf) = 13) And _
           (InStr(1, "0038/0064/0066", Format(pobjstcCobert.Cod_Cobert, "0000"), vbTextCompare) > 0) Then
            pobjstcCobert.Per_Franq = 15
            pobjstcCobert.Val_Franq_Min = 2000
        End If
    End If
    
End Sub

Public Function gfImportarTransmissao(pNossoNumero As String, _
                                      pCodUnicidade As Long, _
                                      pNovNosNumero As String, _
                                      pNovNumPedido As String, _
                                      pNumeroPI As Long, _
                                      pDatRecebimento As Long, _
                                      ByRef pMensagem As String) As Boolean


    'Função: Rotina que transfere os registros das Tabelas de FTP para base quente
    'Parâmetros de entrada...pAntNumItem....Anterior: número do item.
    '                        pmfNovNumItem....Novo: número do item.
    'Parâmetro de saída......pMensagem......Mensagem.

    Dim bdBasDados  As ADODB.Connection 'Base de dados.
    Dim flTabela    As ADODB.Field      'Campos da tabela.
    Dim rsTabela    As ADODB.Recordset  'Registro da tabela.
    Dim lCount      As Long             'Utilizado no for...next.
    Dim lIdeTabela  As String           'Nome da tabela.
    Dim lContinuar  As Boolean          'Indicador se continua o procedimento ou não.
    Dim Dummy       As Variant
    Dim lSelect     As String
    Dim lInsert     As String
    
    On Error GoTo ControleErro
    
    gfImportarTransmissao = False

    '1. Obtém registros da tabela.
    If gfAbrBasDados("P00MPED", gTipP00MPED, gEndP00MPED, bdBasDados, _
            pMensagem) = False Then Exit Function
    
    For lCount = 1 To 18
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
        End Select
        lSelect = "SELECT * FROM FTP_" & lIdeTabela & " WHERE Nosso_Numero = '" & pNossoNumero & "'"
        lSelect = lSelect & " AND COD_UNICIDADE = " & pCodUnicidade
        If gfObtRegistro(bdBasDados, lSelect, rsTabela, pMensagem) = False Then
            Call gpFecha2(bdBasDados)
            Exit Function
        End If
        lContinuar = True
        If rsTabela.EOF = True Then
            lContinuar = False
            If lIdeTabela = "TAB_PED" Then
                pMensagem = "FMP0001 - " & gPrefixo & "Pedido não localizado (" & pNossoNumero & _
                                 " - cópia da cotação/proposta.)."
                Exit Function
            End If
        Else
            If lIdeTabela = "TAB_PED" Then
                If rsTabela("Cod_Tip_Doc") = 9 Then
                    If MsgBox("FMP0002 - " & gPrefixo & " - Pedido já importado! Deseja importar novamente?", vbYesNo) = vbNo Then
                        gfImportarTransmissao = True
                        Exit Function
                    End If
                End If
                'Abrindo uma Transação
                If gfExeSQL(bdBasDados, "Begin Transaction ", pMensagem) <> 0 Then
                    Call gpFecha3(rsTabela, bdBasDados)
                    Exit Function
                End If
            End If
        End If
        If lContinuar = True Then
            'Vai atualizar o Cod_tip_Doc para marcar que já foi carregado
            If lIdeTabela = "TAB_PED" Then
                lSelect = "Update FTP_" & lIdeTabela & " Set Cod_Tip_Doc = 9  WHERE Nosso_Numero = '" & pNossoNumero & "'"
                lSelect = lSelect & " AND COD_UNICIDADE = " & pCodUnicidade
                If gfExeSQL(bdBasDados, lSelect, pMensagem) <> 0 Then
                    GoTo ErrRollBack
                    Exit Function
                End If
            End If
            '2. Grava nova cotação/proposta.
            While Not rsTabela.EOF
                lInsert = "INSERT INTO " & lIdeTabela & " ("
                For Each flTabela In rsTabela.Fields
                    Select Case UCase(flTabela.Name)
                        Case "VAL_RENDAFATU"
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
                        Case "VAL_RENDAFATU"
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
                                If gfPreenchido(flTabela.Value) Then
                                    lInsert = lInsert & "'" & gfForAspas(flTabela.Value) & "', "
                                Else
                                    lInsert = lInsert & "' ', "
                                End If
                            End If
                        Case "NUM_PI"
                            lInsert = lInsert & pNumeroPI & ", "                             'Número do pedido.
                        Case "NOSSO_NUMERO"
                            lInsert = lInsert & "'" & pNovNosNumero & "', "                   'Novo Nosso Número.
                        Case "NUM_PED", "NUM_PROP"
                            lInsert = lInsert & pNovNumPedido & ", "                       'Novo Nosso Número.
                        Case "SIT_DOCUMENTO"
                            lInsert = lInsert & "1, "       'DOCUMENTO NÃO CALCULADO.       'Sit_Documento.
                        Case "DAT_LIBER", "DAT_PRI_CALCULO"
                            lInsert = lInsert & Left(gfForData(3), 8) & ", "                        'Dat_Alt.
                        Case "COD_USER_LIBER"
                            lInsert = lInsert & "'COPIA',"                                  'Cod_User_Alt.
                        Case "DAT_DIGIT"
                            lInsert = lInsert & Left(gfForData(3), 8) & ","
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
                            Select Case flTabela.Type
                                Case adBigInt, adBinary, adInteger, adLongVarBinary, adSmallInt, _
                                     adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, _
                                     adUnsignedTinyInt                          'Numérico - inteiro.
                                    If IsNull(flTabela.Value) Then
                                        lInsert = lInsert & "0, "
                                    Else
                                        lInsert = lInsert & flTabela.Value & ", "
                                    End If
                                Case adCurrency, adSingle, adDouble, adNumeric            'Numérico - com vírgula.
                                    If IsNull(flTabela.Value) Then
                                        lInsert = lInsert & "0, "
                                    Else
                                        lInsert = lInsert & gfForValor(1, flTabela.Value) & ", "
                                    End If
                                Case adChar, adVarChar, adLongVarChar                        'Texto.
                                    If gfPreenchido(flTabela.Value) Then
                                        lInsert = lInsert & "'" & gfForAspas(flTabela.Value) & "', "
                                    Else
                                        lInsert = lInsert & "' ', "
                                    End If
                                Case adBoolean                                  'Sim/Não.
                                    If IsNull(flTabela.Value) Then
                                        lInsert = lInsert & "TRUE, "
                                    Else
                                        lInsert = lInsert & flTabela.Value & ", "
                                    End If
                                Case Else
                                    Call gpFecha3(rsTabela, bdBasDados)
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
                If gfExeSQL(bdBasDados, lInsert, pMensagem) <> 0 Then
                    Call gpFecha3(rsTabela, bdBasDados)
                    GoTo ErrRollBack
                End If
                rsTabela.MoveNext
            Wend
        End If
        Call gpFecha1(rsTabela)
    Next lCount
    lInsert = " Insert into Tab_Ped_Num_Copia values ("
    lInsert = lInsert & "'" & pNovNosNumero & "',"
    lInsert = lInsert & "'" & pNossoNumero & "',"
    lInsert = lInsert & Left(gfForData(3), 8) & ")"
    
    Dummy = gfExeSQL(bdBasDados, lInsert, pMensagem)
    
    Dummy = gfExeSQL(bdBasDados, "Commit Transaction ", pMensagem)
    gfImportarTransmissao = True
    Exit Function
ErrRollBack:

    Call gpGraLog(0, "gfImportarTransmissao ERRO Back - " & pMensagem)
    Dummy = gfExeSQL(bdBasDados, "Rollback Transaction ", pMensagem)
    Exit Function
    
ControleErro:
    Call gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gpGraLog(0, "gfImportarTransmissao ERRO - " & Err & " " & Err.Description)
    
    Resume Next
End Function

Public Function gfSegNovo(Optional ByRef pNossoNumero As String, Optional pMensagem As String, Optional pCod_Ramo As Integer) As Boolean
    'Função: cria cotação - seguro novo.

    Dim bdP00MPED      As ADODB.Connection 'Base de dados: P00MPED.
    Dim lCodProduto     As Integer          'Código do produto.
    Dim lCodProtocolo   As Long             'Código do protocolo.
    Dim lCodSequencia   As Long             'Código da seqüência.
    Dim lDesTarifa      As String           'Descrição da tarifa.
    'Dim pmensagem       As String           'Mensagem.
    Dim lNosNumero      As String           'Nosso número.
    Dim lNumPedido      As String           'Número do pedido.
    Dim lVigDatIni      As Double           'Data de vigência: início.
    Dim lVigDatTer      As Double           'Data de vigência: término.

    gfSegNovo = False
    On Error GoTo ControleErro
    
    '1. Abre base de dados.
    If gfAbrBasDados("P00MPED", gTipP00MPED, gEndP00MPED, bdP00MPED, pMensagem) = False Then
        Call gpGraLog(2, pMensagem)
        Exit Function
        'End
    End If

    '2. Prepara campos.
    'a) Nosso número.
    If pNossoNumero = "" Then
        If gfCalNosNumero(lNosNumero, pMensagem) = False Then
            Call gpGraLog(2, pMensagem)
            Exit Function
            'End
        End If
        pNossoNumero = lNosNumero
    Else
        lNosNumero = pNossoNumero
    End If

    'b) Número do pedido.
    If gfCalNumPedido(lNumPedido, pMensagem, bdP00MPED) = False Then
        Call gpGraLog(2, pMensagem)
        Exit Function
    End If
    gPrefixo = "Cotação " & lNumPedido & ". "
    
    'c) Período de vigência.
    lVigDatIni = gHoje
    If Right(lVigDatIni, 4) = "0229" Then
        lVigDatIni = gHoje - 1
    End If
    lVigDatTer = lVigDatIni + 10000
    
    If Not IsNumeric(pCod_Ramo) Then pCod_Ramo = 0
    
    'd) Produto, protocolo e seqüência.
    If gfBusProduto(lVigDatIni, lCodProduto, lCodProtocolo, lCodSequencia, lDesTarifa, pMensagem, pCod_Ramo) = _
       False Then
        Call gpGraLog(1, pMensagem)
        Exit Function
    End If
    
    '3. Grava registro de dados comuns.
    If gfGraNovPedido(bdP00MPED, lNosNumero, lNumPedido, lVigDatIni, lVigDatTer, _
                   lCodProduto, lCodProtocolo, lCodSequencia, pMensagem, pCod_Ramo) = False Then
        Call gpGraLog(1, pMensagem)
        Exit Function
    End If

    '4. Grava registro de item.
    If gfGraItem(bdP00MPED, lNosNumero, lNumPedido, "0001", pMensagem, 100, , , , , , pCod_Ramo, lCodProduto) = False Then
        Call gpGraLog(1, pMensagem)
        Exit Function
    End If

    '5. Grava registro de comunicação.
    If gfGraRegComunicacao(lNosNumero, pMensagem, bdP00MPED) = False Then
        Call gpGraLog(2, pMensagem)
        Exit Function
    End If

    gfSegNovo = True
    Exit Function
ControleErro:
    Call gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gpGraLog(0, "gfSegNovo ERRO - " & Err & " " & Err.Description & " " & pMensagem)
    
    
    Resume Next
End Function
Function mfCarregaTriagem(ByRef pobjstcPed As stcA46V702B, ByRef pMensagem As String) As Boolean
    Dim lbdBasDados          As ADODB.Connection 'Base de dados.
    Dim lrsTabela            As ADODB.Recordset  'Registro da tabela.
    Dim lSql                 As String
    Dim lI                   As Integer
    Dim lNum_pi              As Long
    Dim lNosNumero           As String
    Dim lPropTabela          As String
    
    On Error GoTo mfCarregaTriagem_Err
    
    Call gfObtSYASMULTINI(lPropTabela)

    mfCarregaTriagem = False

    lNum_pi = pobjstcPed.Num_PI
    lNosNumero = pobjstcPed.Nosso_Numero

    If Val(lNum_pi) = 0 Then
        mfCarregaTriagem = True
        Exit Function
    End If
    
    If gfAbrBasDados("GED", gTipGed, gEndGed, lbdBasDados, pMensagem) = False Then Exit Function
    
    lSql = "Select * "
    lSql = lSql & " from tab_triagem  where num_pi = " & lNum_pi

    If gfObtRegistro(lbdBasDados, lSql, lrsTabela, pMensagem) = False Then
        Call gpFecha2(lbdBasDados)
        Exit Function
    End If

    If lrsTabela.EOF Then
        mfCarregaTriagem = True
        GoTo Finalizar
    End If
    
    pobjstcPed.Nosso_Numero = lNosNumero
    pobjstcPed.Num_Apol_An = Check(lrsTabela("num_apol_ant"))
    pobjstcPed.Dat_Ini_Vig = Check(lrsTabela("dat_ini_vig"))
    pobjstcPed.Dat_Ter_Vig = Check(lrsTabela("dat_fim_vig"))
    pobjstcPed.Cod_Unid_Prod = Check(lrsTabela("cod_unidade"))
    pobjstcPed.Cod_Prod = Check(lrsTabela("cod_produtor"))
    pobjstcPed.Cod_Ramo = Check(lrsTabela("Cod_Ramo"))
    
    pobjstcPed.objstcCtrlEmis.Num_Bloq = Check(lrsTabela("num_bloq"))
    If Check(lrsTabela("num_bloq")) <> 0 Then
        pobjstcPed.objstcCtrlEmis.Cod_Iden_Pag_Antec = e_SimNao_Sim
    End If
    
    If Check(lrsTabela("cod_bco_car")) <> 0 And pobjstcPed.Cod_Bco_Cobr = 0 Then
        pobjstcPed.Cod_Bco_Cobr = Check(lrsTabela("cod_bco_deb"))
    End If
    
    If Check(lrsTabela("cod_bco_deb")) <> 0 And pobjstcPed.Cod_Bco_Cobr = 0 Then
        pobjstcPed.Cod_Bco_Cobr = Check(lrsTabela("cod_bco_deb"))
        pobjstcPed.Cod_Agen_Bas_Cor = Check(lrsTabela("cod_agenc"))
        pobjstcPed.Num_Cta_Bas_Cor = Left(Format(Check(lrsTabela("cod_conta")), "00000000"), 7)
    End If
    
    pobjstcPed.objstcCtrlEmis.Tip_Cotac = Check(lrsTabela("tip_cotac"))
    If pobjstcPed.objstcCtrlEmis.Tip_Cotac = "2" Then
        pobjstcPed.objstcCtrlEmis.Num_Cotac = Check(lrsTabela("num_siscota"))
    End If
    pobjstcPed.objstcCtrlEmis.Dat_Ini_Vig = Check(lrsTabela("dat_ini_vig"))
    pobjstcPed.objstcCtrlEmis.Dat_Ter_Vig = Check(lrsTabela("dat_fim_vig"))
    pobjstcPed.objstcCtrlEmis.Cod_Corr = Check(lrsTabela("cod_corr"))
    pobjstcPed.objstcCtrlEmis.Cod_Corr_2 = Check(lrsTabela("cod_corr2"))
    pobjstcPed.objstcCtrlEmis.Cod_Corr_3 = Check(lrsTabela("cod_corr3"))
    pobjstcPed.objstcCtrlEmis.COD_UNID = Check(lrsTabela("cod_unidade"))
    pobjstcPed.objstcCtrlEmis.Cod_Prod = Check(lrsTabela("cod_produtor"))
    pobjstcPed.objstcCtrlEmis.Cod_Ramo = Check(lrsTabela("Cod_Ramo"))
    
    Dim ObjStcCorretor As stcA46V705B      'Objeto da classe de estrutura stcA46V705B (TAB_PED_CORR)
    Dim ObjCorretor    As clsA46V705A      'Objeto da classe de estrutura stcA46V705B (TAB_PED_CORR)
    
    Set ObjCorretor = New clsA46V705A
    
    If pobjstcPed.ColecaoPedCorr.Count = 0 Then
        lI = 0
        If gfPreenchido(lrsTabela("cod_corr")) Then
            Set ObjStcCorretor = New stcA46V705B
            ObjStcCorretor.Cod_Corr = lrsTabela("cod_corr")
            ObjStcCorretor.Nosso_Numero = pobjstcPed.Nosso_Numero
            ObjStcCorretor.Num_Ordem_Corr = lI + 1
            ObjStcCorretor.Cod_Ramo = pobjstcPed.Cod_Ramo
            ObjStcCorretor.Num_Ped = pobjstcPed.Num_Ped
            If Not ObjCorretor.mfIncluir(ObjStcCorretor, lPropTabela, "") Then
                pMensagem = ObjCorretor.UltimoErro
            End If
            Set ObjStcCorretor = Nothing
            lI = lI + 1
        End If
        If gfPreenchido(lrsTabela("cod_corr2")) Then
            Set ObjStcCorretor = New stcA46V705B
            ObjStcCorretor.Cod_Corr = lrsTabela("cod_corr2")
            ObjStcCorretor.Nosso_Numero = pobjstcPed.Nosso_Numero
            ObjStcCorretor.Num_Ordem_Corr = lI + 1
            ObjStcCorretor.Cod_Ramo = pobjstcPed.Cod_Ramo
            ObjStcCorretor.Num_Ped = pobjstcPed.Num_Ped
            If Not ObjCorretor.mfIncluir(ObjStcCorretor, lPropTabela, "") Then
                pMensagem = ObjCorretor.UltimoErro
            End If
            lI = lI + 1
            Set ObjStcCorretor = Nothing
        End If
        If gfPreenchido(lrsTabela("cod_corr3")) Then
            Set ObjStcCorretor = New stcA46V705B
            ObjStcCorretor.Cod_Corr = lrsTabela("cod_corr3")
            ObjStcCorretor.Nosso_Numero = pobjstcPed.Nosso_Numero
            ObjStcCorretor.Num_Ordem_Corr = lI + 1
            ObjStcCorretor.Cod_Ramo = pobjstcPed.Cod_Ramo
            ObjStcCorretor.Num_Ped = pobjstcPed.Num_Ped
            If Not ObjCorretor.mfIncluir(ObjStcCorretor, lPropTabela, "") Then
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
    
Finalizar:

    Exit Function
    
mfCarregaTriagem_Err:
    MsgBox "Erro no módulo de "
End Function

Sub gpAtualizaIni()
Dim lPathPrinc As String
Dim lPathArquivos As String
Dim Aux           As String
Dim lMantem       As Boolean

If gTipUsuario = e_Tipo_Usu_Yasuda Then Exit Sub

Aux = ""
Aux = LerIni("MULTPROB_V0", "NOME", gAppPath & "SyasMult.INI")
lMantem = True
If Trim(Aux) <> "" Then
    lPathArquivos = ""
    If InStrRev(Aux, "\") > 0 Then
        lPathArquivos = Left(Aux, InStrRev(Aux, "\"))
    End If
    If gfPreenchido(lPathArquivos) = True And (gTipUsuario = e_Tipo_Usu_Corretor Or gTipUsuario = e_Tipo_Usu_CorretorYasuda) Then
        'Vai corrigir o path alterado indevidamente na versão de 28/09/2007
        lMantem = False
    Else
        Call gpGraLog(0, "gpAtualizaIni: saiu sem reescrever")
        Exit Sub  'Se já existir desconsidera
    End If
End If

lPathPrinc = LerIni("MULTCOT1", "CONFIG", gAppPath & "SyasMult.INI")
lPathArquivos = ""
If InStrRev(lPathPrinc, "\") > 0 Then
    lPathArquivos = Left(lPathPrinc, InStrRev(lPathPrinc, "\"))
End If

Call gpGraLog(0, "gpAtualizaIni: Caminho encontrado = " & lPathArquivos)
If lMantem = False Then
    lPathArquivos = ""
    Call gpGraLog(0, "gpAtualizaIni: corrigindo caminho = " & lPathArquivos)
End If
Call gpGraLog(0, "gpAtualizaIni: Atualizando caminho CONFIG =" & lPathArquivos)

Call GravarINI("MULTCOT1", "CONFIG", lPathArquivos & "MULTCOT1.CXY", gAppPath & "SyasMult.INI")
Call GravarINI("MULTCOT1", "IMAGEM", lPathArquivos & "MULTCOT1.WMF", gAppPath & "SyasMult.INI")

Call GravarINI("MULTCOT2", "CONFIG", lPathArquivos & "MULTCOT2.CXY", gAppPath & "SyasMult.INI")
Call GravarINI("MULTCOT2", "IMAGEM", lPathArquivos & "MULTCOT2.WMF", gAppPath & "SyasMult.INI")

Call GravarINI("MULTCOT3", "CONFIG", lPathArquivos & "MULTCOT3.CXY", gAppPath & "SyasMult.INI")
Call GravarINI("MULTCOT3", "IMAGEM", lPathArquivos & "MULTCOT3.WMF", gAppPath & "SyasMult.INI")

Call GravarINI("MULTCOTA", "CONFIG", lPathArquivos & "MULTCOTA.CXY", gAppPath & "SyasMult.INI")
Call GravarINI("MULTCOTA", "IMAGEM", lPathArquivos & "MULTCOTA.WMF", gAppPath & "SyasMult.INI")

Call GravarINI("MULTCOTC", "CONFIG", lPathArquivos & "MULTCOTC.CXY", gAppPath & "SyasMult.INI")
Call GravarINI("MULTCOTC", "IMAGEM", lPathArquivos & "MULTCOTC.WMF", gAppPath & "SyasMult.INI")

Call GravarINI("MULTPRO1", "CONFIG", lPathArquivos & "MULTPRO1.CXY", gAppPath & "SyasMult.INI")
Call GravarINI("MULTPRO1", "IMAGEM", lPathArquivos & "MULTPRO1.WMF", gAppPath & "SyasMult.INI")

Call GravarINI("MULTPRO2", "CONFIG", lPathArquivos & "MULTPRO2.CXY", gAppPath & "SyasMult.INI")
Call GravarINI("MULTPRO2", "IMAGEM", lPathArquivos & "MULTPRO2.WMF", gAppPath & "SyasMult.INI")

Call GravarINI("MULTPRO3", "CONFIG", lPathArquivos & "MULTPRO3.CXY", gAppPath & "SyasMult.INI")
Call GravarINI("MULTPRO3", "IMAGEM", lPathArquivos & "MULTPRO3.WMF", gAppPath & "SyasMult.INI")

Call GravarINI("MULTPRO4", "CONFIG", lPathArquivos & "MULTPRO4.CXY", gAppPath & "SyasMult.INI")
Call GravarINI("MULTPRO4", "IMAGEM", lPathArquivos & "MULTPRO4.WMF", gAppPath & "SyasMult.INI")

Call GravarINI("MULTPROB", "CONFIG", lPathArquivos & "MULTPROB.CXY", gAppPath & "SyasMult.INI")
Call GravarINI("MULTPROB", "IMAGEM", lPathArquivos & "MULTPROB.WMF", gAppPath & "SyasMult.INI")

Call GravarINI("MULTPROC", "CONFIG", lPathArquivos & "MULTPROC.CXY", gAppPath & "SyasMult.INI")
Call GravarINI("MULTPROC", "IMAGEM", lPathArquivos & "MULTPROC.WMF", gAppPath & "SyasMult.INI")

Call GravarINI("MULTPROD", "CONFIG", lPathArquivos & "MULTPROD.CXY", gAppPath & "SyasMult.INI")
Call GravarINI("MULTPROD", "IMAGEM", lPathArquivos & "MULTPROD.WMF", gAppPath & "SyasMult.INI")


Call GravarINI("MULTCOT1_V0", "NOME", lPathArquivos & "MULTCOT1_V0", gAppPath & "SyasMult.INI")
Call GravarINI("MULTCOT2_V0", "NOME", lPathArquivos & "MULTCOT2_V0", gAppPath & "SyasMult.INI")
Call GravarINI("MULTCOT3_V0", "NOME", lPathArquivos & "MULTCOT3_V0", gAppPath & "SyasMult.INI")
Call GravarINI("MULTPRO1_V0", "NOME", lPathArquivos & "MULTPRO1_V0", gAppPath & "SyasMult.INI")
Call GravarINI("MULTPRO2_V0", "NOME", lPathArquivos & "MULTPRO2_V0", gAppPath & "SyasMult.INI")
Call GravarINI("MULTPRO3_V0", "NOME", lPathArquivos & "MULTPRO3_V0", gAppPath & "SyasMult.INI")
Call GravarINI("MULTPRO4_V0", "NOME", lPathArquivos & "MULTPRO4_V0", gAppPath & "SyasMult.INI")
Call GravarINI("MULTPROB_V0", "NOME", lPathArquivos & "MULTPROB_V0", gAppPath & "SyasMult.INI")

End Sub





Private Function mfGraRenItem_Roubo(prstItem As ADODB.Recordset, pNosNumero As String, _
                    pBdP00APOL As ADODB.Connection, pMensagem As String, _
                    pIndSinistro As e_SimNao, pNumPed As Long, pCod_Dizer As Integer) As Boolean
    'Função: Grava os registros de comum com os dados da apólice (Renovação Yasuda).
    'Parâmetros de entrada :prstItem.........Registro de item(apólice)
    '                       pNosNumero.......Nosso número.
    '                       pBdP00APOL.......Arquivo de conexão
    'Parâmetro de saída:    pMensagem........Mensagem de erro.
    
    Dim lSql            As String                   'Consulta SQL.
    Dim rsAPOL          As ADODB.Recordset          'Registro.

    '[FRH]->
    'Conferir coberturas na tabela POOMULTI:TAB_COB_RELACAO
    Dim lSql_buscaCob As String
    Dim rs_buscaCob As ADODB.Recordset
    Dim bdP00MULTI As ADODB.Connection
    '<-[FRH]
    
    Dim lObjPedLoc      As clsA46V708A              'Objeto da classe clsA46V708 (TAB_PED_LOC).
    Dim lobjstcPedLoc   As stcA46V708B              'Objeto da classe de estrutura stcA46V708B (TAB_PED_LOC).
    Dim lObjClau        As clsA46V703A              'Objeto da classe clsA46V703A (TAB_PED_CLAU).
    Dim lobjstcClau     As stcA46V703B              'Objeto da classe de estrutura stcA46V703B (TAB_PED_CLAU).
    Dim lobjCobert      As clsA46V704A              'Objeto da classe clsA46V704A (TAB_PED_COBERT).
    Dim lobjstcCobert   As stcA46V704B              'Objeto da classe de estrutura stcA46V704B (TAB_PED_COBERT).
    Dim lobjInsp        As clsA46V717A              'Objeto da classe clsA46V717A (TAB_PED_INSP).
    Dim lObjstcInsp     As stcA46V717B              'Objeto da classe de estrutura stcA46V717B (TAB_PED_INSP).
    Dim lNumero         As String
    Dim lComplemento    As String
    Dim lPos            As Long
    Dim lNomeTab        As String
    Dim lValorRisco     As Double
    Dim lObjPedFator    As clsA46V720A
    Dim lobjstcPedFator As stcA46V720B
    Dim lPropTabela     As String
    
    On Error GoTo ControleErro
    
    Call gfObtSYASMULTINI(lPropTabela)

    If gBaseLocal = True Then
        lNomeTab = "TAB_APOL"
    Else
        lNomeTab = "GED.DBO.TAB_PED"
    End If
    
    Set lObjPedLoc = New clsA46V708A
    Set lobjstcPedLoc = New stcA46V708B
    With lobjstcPedLoc
        .Nosso_Numero = pNosNumero
        .Num_Ped = pNumPed
        If gNum_Pi = 0 Then
            .Num_PI = Right(Year(Date), 2) & Format(pNumPed, "00000000")
        Else
            .Num_PI = gNum_Pi
        End If
        .Cod_Ramo = gfTratarNulo(prstItem("Cod_Ramo"), tpoNumerico)
        .Num_Item = 1
        .Per_Desc_Apar = 0
        .Per_Desc_Progres = 0
        .Per_IS_Vr = 0
        .Val_Coef_IS_Vr = 0
        .Val_Vr = gfTratarNulo(prstItem("Val_IS"), TpoValor)
        lValorRisco = .Val_Vr
        .Tab_Clasf = 0
        .Cod_Clasf = ""
        .Cod_Compl_Clasf = ""
        .Dsc_Clasf = ""
        .Cod_Constr = 0
        .Cod_User_Liber = gNomUsuario6
        .Dat_Liber = Left(gfForData(3), 8)
        .Hor_Liber = Mid(gfForData(3), 9, 4)
        .Cod_Sit_item = e_SitItem_NaoCalc
        .Per_Bonus = 0
        .Num_Apol_Anterior = prstItem("Num_APol")
        .Num_Item_Anterior = 1
        If lobjstcPedLoc.mBonus_Anterior = 0 Then
            .Per_Bonus = lObjPedLoc.mfRetornaBonusApolice(lobjstcPedLoc)
        Else
            .Per_Bonus = lobjstcPedLoc.mBonus_Anterior
        End If
        If pIndSinistro = e_SimNao_Sim Then
            .Per_Bonus = .Per_Bonus - 5
        Else
            .Per_Bonus = .Per_Bonus + 5
        End If
        If .Per_Bonus < 0 Then .Per_Bonus = 0
        If .Per_Bonus > 25 Then .Per_Bonus = 25
        gDesc_Bonus = .Per_Bonus
        .Per_Tot_Protec = .gPer_Tot_Protec
        .Cod_Produto = gCod_Plano
        .Cod_Regiao = gfTratarNulo(prstItem("Cod_Reg"), tpoNumerico)
        If pCod_Dizer = 0 Then
            .Tip_Emissao = e_TipEmissao_RenYasuda
        Else
            .Tip_Emissao = e_TipEmissao_Endosso
        End If
        .Ind_Sinistro = pIndSinistro
        .Cod_Iden_Tarif = gCodProduto

    End With
    
    If Not lObjPedLoc.mfIncluir(lobjstcPedLoc, lPropTabela, "") Then
        pMensagem = lObjPedLoc.UltimoErro
        Exit Function
    End If
    
    'Clausulas
    lSql = " Select * from " & lNomeTab & "_CLAU "
    If gBaseLocal = False And pBdP00APOL Like "*Provider=MSDASQL.*" Then
        lSql = lSql & " inner join GED.DBO.tab_ctrl_emis  as ctrl on " & lNomeTab & "_CLAU.num_apol = ctrl.num_apol"
        lSql = lSql & " and ctrl.num_endo = " & lNomeTab & "_CLAU.num_endo"
        lSql = lSql & " and ctrl.num_PI = " & lNomeTab & "_CLAU.num_PI"
    End If
    lSql = lSql & " Where " & lNomeTab & "_CLAU.Num_Apol = " & prstItem("Num_Apol")
    lSql = lSql & " and " & lNomeTab & "_CLAU.num_item = " & 1
    
    If Not gfObtRegistro(pBdP00APOL, lSql, rsAPOL, pMensagem) Then
        GoTo ControleErro
    End If
    
    Set lObjClau = New clsA46V703A
    Set lObjClau.ConexaoBD = lObjPedLoc.ConexaoBD
    Do While Not rsAPOL.EOF
        Set lobjstcClau = New stcA46V703B
        With lobjstcClau
            .Cod_Ramo = gfTratarNulo(rsAPOL("Cod_Ramo"), tpoNumerico)
            .Nosso_Numero = pNosNumero
            .Num_Ped = pNumPed
            If gNum_Pi = 0 Then
                .Num_PI = Right(Year(Date), 2) & Format(pNumPed, "00000000")
            Else
                .Num_PI = gNum_Pi
            End If
            .Cod_Clau = gfTratarNulo(rsAPOL("Cod_Clau"), tpoNumerico)
            .Num_Item = lobjstcPedLoc.Num_Item
        End With
        If Not lObjClau.mfIncluir(lobjstcClau, lPropTabela, "") Then
            pMensagem = lObjClau.UltimoErro
            Exit Function
        End If
        rsAPOL.MoveNext
    Loop
    
    'COberturas
    lSql = " Select " & lNomeTab & "_COBERT.* from " & lNomeTab & "_COBERT "
    If gBaseLocal = False And pBdP00APOL Like "*Provider=MSDASQL.*" Then
        lSql = lSql & " inner join GED.DBO.tab_ctrl_emis  as ctrl on GED.DBO.TAB_PED_COBERT.num_apol = ctrl.num_apol"
        lSql = lSql & " and ctrl.num_endo = GED.DBO.TAB_PED_COBERT.num_endo"
        lSql = lSql & " and ctrl.num_PI = GED.DBO.TAB_PED_COBERT.num_PI"
    End If
    lSql = lSql & " Where " & lNomeTab & "_COBERT.Num_Apol = " & prstItem("Num_Apol")
    lSql = lSql & " and " & lNomeTab & "_COBERT.num_item = " & 1
    
    If Not gfObtRegistro(pBdP00APOL, lSql, rsAPOL, pMensagem) Then
        GoTo ControleErro
    End If
    
    Set lobjCobert = New clsA46V704A
   'Set lObjCobert.ConexaoBD = lObjPedLoc.ConexaoBD
    '[FRH]->
    If Not gfAbrBasDados("P00MULTI", gTipP00MULTI, gEndP00MULTI, bdP00MULTI, pMensagem) Then
        GoTo ControleErro
    End If
    '<-[FRH]
    Do While Not rsAPOL.EOF
        Set lobjstcCobert = New stcA46V704B
        With lobjstcCobert
            .Cod_Ramo = gfTratarNulo(rsAPOL("Cod_Ramo"), tpoNumerico)
            .Nosso_Numero = pNosNumero
            .Num_Ped = pNumPed
            If gNum_Pi = 0 Then
                .Num_PI = Right(Year(Date), 2) & Format(pNumPed, "00000000")
            Else
                .Num_PI = gNum_Pi
            End If
            .Num_Item = lobjstcPedLoc.Num_Item
                        
            '[FRH]->
            'Rotina para conferir alterações de códigos de cobertura,
            'na *POOMULTI.TAB_COB_RELACAO* existe um de-para de cobs.anteriores e atuais.
            
            '[Valor original Cod_Cobert da TAB_APOL]
            .Cod_Cobert = gfTratarNulo(rsAPOL("cod_cobert"), tpoNumerico)
            
            lSql_buscaCob = "SELECT COD_COB_ATUAL FROM TAB_COB_RELACAO WHERE COD_COB_ANTERIOR=" & gfTratarNulo(rsAPOL("Cod_Cobert"), tpoNumerico)
            If gfObtRegistro(bdP00MULTI, lSql_buscaCob, rs_buscaCob, pMensagem) Then
                If Not (rs_buscaCob.BOF And rs_buscaCob.EOF) Then
                    .Cod_Cobert = gfTratarNulo(rs_buscaCob("cod_cob_atual"), tpoNumerico)
                End If
            End If
            '<-[FRH]
            
       '    .Tab_Cobert = gfTratarNulo(rsApol("Tab_Cobert"), tpoNumerico)
            If gCod_Plano > 1010 Then
                .Tab_Cobert = gCod_Plano
            Else
                .Tab_Cobert = 1
            End If
            
                   
            .Per_Desc_Exp = gDesc_Bonus
                       
            .Val_IS = gfTratarNulo(rsAPOL("Val_IS"), tpoNumerico)
            .Val_Risco_Cob = .Val_IS
            
            If pCod_Dizer <> 0 Then
                If gCod_Plano > 1010 And .Cod_Ramo = e_CodRamo_Residencial Then
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
            Call mpBuscaLimFranquia(lobjstcPedLoc, lobjstcCobert)
            '11/01/2010 - Ficha: 1000003 - franquia e valor de franquia - fim
            
        End With

        If Not lobjCobert.mfIncluir(lobjstcCobert, lPropTabela, "") Then
            If UCase(lobjCobert.UltimoErro) Like "*DUPLICATE*" And Not rs_buscaCob.EOF And Not rs_buscaCob.BOF Then
                GoTo Proximo
            Else
                pMensagem = lobjCobert.UltimoErro
                Exit Function
            End If
        End If
        If lobjstcCobert.Cod_Cobert = 1 And lValorRisco = 0 Then
            lobjstcPedLoc.Val_Vr = lobjstcCobert.Val_IS
            lobjstcPedLoc.Per_IS_Vr = 100
            If Not lObjPedLoc.mfAlterarValorRisco(lobjstcPedLoc, lPropTabela, "") Then
                pMensagem = lObjPedLoc.UltimoErro
                Exit Function
            End If
        End If
Proximo:
        rsAPOL.MoveNext
    Loop
        
    '[FRH]->
    Call gpFecha2(bdP00MULTI)
    '<-[FRH]
    
    Set lobjInsp = New clsA46V717A
    Set lObjstcInsp = New stcA46V717B
    lObjstcInsp.Nosso_Numero = pNosNumero
    lObjstcInsp.Num_Item = lobjstcPedLoc.Num_Item
    lObjstcInsp.Num_Logr_Risco = Val(lNumero)
    'correção
    lObjstcInsp.Cod_Ramo = lobjstcPedLoc.Cod_Ramo
    lObjstcInsp.Num_Ped = lobjstcPedLoc.Num_Ped
    '
    If Not lobjInsp.mfIncluir(lObjstcInsp, lPropTabela, "") Then
        pMensagem = lobjInsp.UltimoErro
        Exit Function
    End If
    
        
    'Ficha 1000409 - Inclusão de Fatores de Proteção em Renovações Yasuda
    If gBaseLocal = True Then
        lNomeTab = "TAB_APOL"
    Else
        lNomeTab = "RamosDiversos.dbo.Tab_Ped"
    End If
    lSql = " Select * from " & lNomeTab & "_FAT_RISCO "
    If gBaseLocal = False And pBdP00APOL Like "*Provider=MSDASQL.*" Then
        lSql = lSql & " inner join GED.DBO.tab_ctrl_emis  as ctrl on " & lNomeTab & "_FAT_RISCO.num_apol = ctrl.num_apol"
        lSql = lSql & " and ctrl.num_endo = " & lNomeTab & "_FAT_RISCO.num_endo"
        lSql = lSql & " and ctrl.num_PI = " & lNomeTab & "_FAT_RISCO.num_PI"
    End If
    lSql = lSql & " Where " & lNomeTab & "_FAT_RISCO.Num_Apol = " & prstItem("Num_Apol")
    lSql = lSql & " and " & lNomeTab & "_FAT_RISCO.num_item = " & 1
'    lSql = lSql & " and " & lNomeTab & "_FAT_RISCO.num_Endo = " & prstItem("NUm_Endo2")
    
    On Error Resume Next
    If Not gfObtRegistro(pBdP00APOL, lSql, rsAPOL, pMensagem) Then
        GoTo ControleErro
    End If
    
    Set lObjPedFator = New clsA46V720A
    Set lObjPedFator.ConexaoBD = lObjPedLoc.ConexaoBD
    Do While Not rsAPOL.EOF
        Set lobjstcPedFator = New stcA46V720B
        With lobjstcPedFator
            .Cod_Ramo = gfTratarNulo(rsAPOL("Cod_Ramo"), tpoNumerico)
            .Nosso_Numero = pNosNumero
            .Num_Ped = pNumPed
            If gNum_Pi = 0 Then
                .Num_PI = Right(Year(Date), 2) & Format(pNumPed, "00000000")
            Else
                .Num_PI = gNum_Pi
            End If
            .Num_Item = lobjstcPedLoc.Num_Item
            .Cod_Fator = gfTratarNulo(rsAPOL("Cod_Fator"), tpoNumerico)
            .Dsc_Compl_Fator = gfTratarNulo(rsAPOL("DSC_COMPL_FATOR"), tpoNumerico)
        End With
        If Not lObjPedFator.mfIncluir(lobjstcPedFator, lPropTabela, "") Then
            pMensagem = lObjPedFator.UltimoErro
            Exit Function
        End If
        rsAPOL.MoveNext
    Loop
    On Error GoTo ControleErro
    
        
    Set lObjClau = Nothing
    Set lobjCobert = Nothing
    Set lObjPedLoc = Nothing
    Set lobjstcClau = Nothing
    Set lobjstcCobert = Nothing
    Set lobjstcPedLoc = Nothing
    
    mfGraRenItem_Roubo = True
    Exit Function
ControleErro:
    Set lObjClau = Nothing
    Set lobjCobert = Nothing
    Set lObjPedLoc = Nothing
    Set lobjstcClau = Nothing
    Set lobjstcCobert = Nothing
    Set lobjstcPedLoc = Nothing
    Call gpGraLog(1, Err & " " & Err.Description, vbCritical)
    Call gpGraLog(0, "gfGraRenItem ERRO - " & Err & " " & Err.Description & " " & pMensagem)
End Function

Public Function gfRetornaCodProdutor(ByVal pCodCorr As Long, ByVal pCodUnidProd As Integer) As Integer
Dim lCN             As ADODB.Connection
Dim lRs             As ADODB.Recordset
Dim lCodRetorno     As Integer
Dim lMsgErr         As String
Dim lSelect         As String

On Error GoTo TrataErro: gfRetornaCodProdutor = 0

    If Not gPrograma = "D046O003" Then GoTo SAIDA
    
    If gfAbrBasDados("P00MPED", gTipP00MPED, gEndP00MPED, lCN, _
                     lMsgErr) = False Then GoTo TrataErro
    
    lSelect = "EXEC [RamosDiversos].[dbo].[proc_RetornaProdutor] " & pCodCorr & ", " & pCodUnidProd

    If Not gfObtRegistro(lCN, lSelect, lRs, lMsgErr) Then
        GoTo TrataErro
    ElseIf lRs.EOF = False Then
        lCodRetorno = gfTratarNulo(lRs("Cod_Prod"), tpoNumerico)
    End If

    gfRetornaCodProdutor = lCodRetorno

SAIDA:
    GoSub FechaObj
    Exit Function
TrataErro:
    If Not gfPreenchido(lMsgErr) Then
        lMsgErr = Err & " - " & Error
    End If
    lMsgErr = "gfRetornaCodProdutor Erro:" & lMsgErr
    Call gpGraLog(0, lMsgErr)
    GoTo SAIDA
FechaObj:
    Call gpFecha3(lRs, lCN)
    Set lRs = Nothing
    Set lCN = Nothing
    Return
End Function

