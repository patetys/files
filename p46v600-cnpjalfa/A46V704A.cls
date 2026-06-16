VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA46V704A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================================================================================

' Class Name: clsA46V704A --> Tabela Tab_Ped_Cobert
' Data/Hora: 31/08/2005 - 11:52:48
' Sistema: Multi
' Thais Juren Gomes

' ======================================================================================================================================================


' ======================================================================================================================================================
Public ConexaoBD As ADODB.Connection    ' Conexão com o Banco de Dados
Public Colecao As Collection    ' Coleção de objetos do tipo stcA46V704B
Public UltimoErro As String           ' Descrição do último erro gerado
Public ColecaoNovaFranquia            ' Coleção de objetos da tab_ped_cobert_franquias

Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "A46V704A"
End Property

Public Function mfAlterarDescontoComercial(ByRef pobjstc As stcA46V704B, _
                                           ByVal pPropTabela As String, _
                                           ByVal pNomeTabImport As String) As Boolean

    Dim lUpdate As String

    On Error GoTo ControleErros

    If Not mfConectar Then GoTo ControleErros

    lUpdate = " Update " & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert set Per_Desc_Coml = " & M46V999.gfFormatarCamposSQL(pobjstc.Per_Desc_Coml)
    lUpdate = lUpdate & " WHERE Nosso_Numero = " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lUpdate = M46V999.gfWhereAnd(lUpdate) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If
    lUpdate = lUpdate & " and NUM_ITEM       = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item)
    lUpdate = lUpdate & " and COD_COBERT     = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Cobert)


    If M46V999.gfExeSQL(ConexaoBD, lUpdate, UltimoErro) <> 0 Then GoTo ControleErros

    mfAlterarDescontoComercial = True

    Exit Function
ControleErros:
    mfAlterarDescontoComercial = False
    UltimoErro = "A46V704A.mfAlterarDescontoComercial ERRO - " & Err & " " & Err.Description & _
                 vbCrLf & " SQL: " & lUpdate
End Function



Public Function mfAlterarValoresPrecificacao(ByRef pobjstc As stcA46V704B, _
                                             ByVal pPropTabela As String, _
                                             ByVal pNomeTabImport As String) As Boolean

    Dim lUpdate As String
    Dim lRegAlterados As Long

    On Error GoTo ControleErros

    If Not mfConectar Then GoTo ControleErros

    lUpdate = " Update " & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert set Per_Desc_Coml = " & M46V999.gfFormatarCamposSQL(pobjstc.Per_Desc_Coml)
    lUpdate = lUpdate & ", Val_Pr = " & M46V999.gfFormatarCamposSQL(pobjstc.VAL_PR)
    lUpdate = lUpdate & ", Val_Pr_Anu  = " & M46V999.gfFormatarCamposSQL(pobjstc.Val_Pr_Anu)
    lUpdate = lUpdate & ", Val_Pr_Tarif_Anu  = " & M46V999.gfFormatarCamposSQL(pobjstc.Val_Pr_Tarif_Anu)
    lUpdate = lUpdate & ", Val_Pr_Tarif  = " & M46V999.gfFormatarCamposSQL(pobjstc.Val_Pr_Tarif)
    lUpdate = lUpdate & ", Val_Pr_Resg_Anu  = " & M46V999.gfFormatarCamposSQL(pobjstc.Val_Pr_Resg_Anu)
    lUpdate = lUpdate & ", Val_Pr_Resg  = " & M46V999.gfFormatarCamposSQL(pobjstc.Val_Pr_Resg)
    lUpdate = lUpdate & ", Per_Taxa = " & M46V999.gfFormatarCamposSQL(pobjstc.Per_Taxa)
    lUpdate = lUpdate & ", Per_Franq = " & M46V999.gfFormatarCamposSQL(pobjstc.Per_Franq)
    lUpdate = lUpdate & ", Val_Franq_Min = " & M46V999.gfFormatarCamposSQL(pobjstc.Val_Franq_Min)
    If pobjstc.ObjPai.Cod_Ramo <> e_CodRamo_Roubos Then
        lUpdate = lUpdate & ", Tab_Cobert = " & M46V999.gfFormatarCamposSQL(pobjstc.ObjPai.Tab_Clasf)
    End If
    lUpdate = lUpdate & " WHERE Nosso_Numero = " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lUpdate = M46V999.gfWhereAnd(lUpdate) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If
    lUpdate = lUpdate & " and NUM_ITEM       = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item)
    lUpdate = lUpdate & " and COD_COBERT     = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Cobert)

    If M46V999.gfExeSQL(ConexaoBD, lUpdate, UltimoErro) <> 0 Then GoTo ControleErros

    lUpdate = " Update " & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert_Compl set Cod_Classe = " & M46V999.gfFormatarCamposSQL(pobjstc.COD_CLASSE)
    lUpdate = lUpdate & ", Alterar_Cobert = " & M46V999.gfFormatarCamposSQL(pobjstc.Alterar_Cobert)
    lUpdate = lUpdate & ", Formula = " & M46V999.gfFormatarCamposSQL(pobjstc.Formula)
    If pobjstc.ObjPai.Cod_Ramo <> e_CodRamo_Roubos Then
        lUpdate = lUpdate & ", Tab_Cobert = " & M46V999.gfFormatarCamposSQL(pobjstc.ObjPai.Tab_Clasf)
    End If
    lUpdate = lUpdate & ", Val_PR_Orig = " & M46V999.gfFormatarCamposSQL(pobjstc.Val_Pr_Orig)
    lUpdate = lUpdate & ", PremioInformado = " & M46V999.gfFormatarCamposSQL(pobjstc.PremioInformado)
    lUpdate = lUpdate & ", VAL_PR_BATENTE  = " & M46V999.gfFormatarCamposSQL(pobjstc.VAL_PR_BATENTE)
    lUpdate = lUpdate & ", Per_Desc_Franq  = " & M46V999.gfFormatarCamposSQL(pobjstc.Per_Desc_Franq)

    ''2016-10-19~2018-02-28: VerbaCom
    If UCase(pPropTabela) Like "*SYASPORTAL*" Or UCase(pPropTabela) Like "*RAMOSDIVERSOS*" Then
        lUpdate = lUpdate & ",Val_VerbaComProp = " & M46V999.gfFormatarCamposSQL(pobjstc.Val_VerbaComProp)
    End If
    
    lUpdate = lUpdate & " WHERE Nosso_Numero = " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lUpdate = M46V999.gfWhereAnd(lUpdate) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If
    lUpdate = lUpdate & " and NUM_ITEM       = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item)
    lUpdate = lUpdate & " and COD_COBERT     = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Cobert)

    If M46V999.gfExeSQL(ConexaoBD, lUpdate, UltimoErro) <> 0 Then GoTo ControleErros

    mfAlterarValoresPrecificacao = True

    Exit Function
ControleErros:
    mfAlterarValoresPrecificacao = False
    UltimoErro = "A46V704A.mfAlterarValoresPrecificacao ERRO - " & Err & " " & Err.Description & _
                 vbCrLf & " SQL: " & lUpdate
End Function

Public Function mfAlterarValoresPremio(ByRef pobjstc As stcA46V704B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As Boolean

    Dim lUpdate As String
    Dim lRegAlterados As Long

    On Error GoTo ControleErros

    If Not mfConectar Then GoTo ControleErros

    lUpdate = " Update " & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert set "
    lUpdate = lUpdate & " Val_Pr = " & M46V999.gfFormatarCamposSQL(pobjstc.VAL_PR)
    lUpdate = lUpdate & ", Val_Pr_Tarif  = " & M46V999.gfFormatarCamposSQL(pobjstc.Val_Pr_Tarif)
    lUpdate = lUpdate & ", Val_Pr_Resg  = " & M46V999.gfFormatarCamposSQL(pobjstc.Val_Pr_Resg)
    If pobjstc.ObjPai.Cod_Ramo <> e_CodRamo_Roubos Then
        lUpdate = lUpdate & ", Tab_Cobert = " & M46V999.gfFormatarCamposSQL(pobjstc.ObjPai.Tab_Clasf)
    End If
    lUpdate = lUpdate & " WHERE Nosso_Numero = " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lUpdate = M46V999.gfWhereAnd(lUpdate) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If
    lUpdate = lUpdate & " and NUM_ITEM       = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item)
    lUpdate = lUpdate & " and COD_COBERT     = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Cobert)

    If M46V999.gfExeSQL(ConexaoBD, lUpdate, UltimoErro) <> 0 Then GoTo ControleErros

    mfAlterarValoresPremio = True

    Exit Function
ControleErros:
    mfAlterarValoresPremio = False
    UltimoErro = "A46V704A.mfAlterarValoresPremio ERRO - " & Err & " " & Err.Description & _
                 vbCrLf & " SQL: " & lUpdate
End Function


Private Function mfFormatarInclusaoSQL(ByRef pobjstc As stcA46V704B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String, _
                                       Optional ByVal CodTab As Integer) As String
    Dim lInsert As String                'Query.

    If CodTab = 2 Then GoTo Tab_Ped_Cobert_Compl

    lInsert = " INSERT INTO " & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert ( "
    lInsert = lInsert & " Nosso_numero,"
    If pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_" Then
        lInsert = lInsert & " Cod_Unicidade,"
    End If
    lInsert = lInsert & " Cod_Cobert,"
    lInsert = lInsert & " Num_Ped,"
    lInsert = lInsert & " Cod_Org,"
    lInsert = lInsert & " Cod_Ramo,"
    lInsert = lInsert & " Num_Apol,"
    lInsert = lInsert & " Num_Endo,"
    lInsert = lInsert & " Num_item,"
    lInsert = lInsert & " Tab_Cobert,"
    lInsert = lInsert & " Val_IS,"
    lInsert = lInsert & " Per_Taxa,"
    lInsert = lInsert & " Per_Desc_Coml,"
    lInsert = lInsert & " Per_Desc_Exp,"
    lInsert = lInsert & " Val_PR,"
    lInsert = lInsert & " Val_Pr_Anu,"
    lInsert = lInsert & " Val_Pr_Tarif,"
    lInsert = lInsert & " Val_Pr_Tarif_Anu,"
    lInsert = lInsert & " Val_Pr_Resg,"
    lInsert = lInsert & " Val_Pr_Resg_Anu,"
    lInsert = lInsert & " Per_Franq,"
    lInsert = lInsert & " Val_Franq_Min,"
    lInsert = lInsert & " Val_Franq_Max,"
    lInsert = lInsert & " Num_PI"
    lInsert = lInsert & ") VALUES ("
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero) & ","
    If pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_" Then
        lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade) & ","
    End If
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Cobert) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Ped) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Org) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Ramo) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Apol) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Endo) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Tab_Cobert) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Val_IS) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Per_Taxa) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Per_Desc_Coml) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Per_Desc_Exp) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.VAL_PR) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Val_Pr_Anu) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Val_Pr_Tarif) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Val_Pr_Tarif_Anu) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Val_Pr_Resg) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Val_Pr_Resg_Anu) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Per_Franq) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Val_Franq_Min) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Val_Franq_Max) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.NUM_PI) & ") "

    mfFormatarInclusaoSQL = lInsert
    Exit Function

Tab_Ped_Cobert_Compl:

    lInsert = " INSERT INTO " & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert_Compl  ( "
    lInsert = lInsert & " Nosso_numero,"
    If pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_" Then
        lInsert = lInsert & " Cod_Unicidade,"
    End If
    lInsert = lInsert & " Cod_Cobert,"
    lInsert = lInsert & " Num_Ped,"
    lInsert = lInsert & " Cod_Org,"
    lInsert = lInsert & " Cod_Ramo,"
    lInsert = lInsert & " Num_Apol,"
    lInsert = lInsert & " Num_Endo,"
    lInsert = lInsert & " Num_item,"
    lInsert = lInsert & " Tab_Cobert,"
    lInsert = lInsert & " Num_Meses_PI,"
    lInsert = lInsert & " Formula,"
    lInsert = lInsert & " Cod_Classe,"
    lInsert = lInsert & " Val_Fator_Franq,"
    lInsert = lInsert & " Alterar_Cobert,"
    lInsert = lInsert & " Val_Risco_Cob,"
    lInsert = lInsert & " Num_PI, "
    lInsert = lInsert & " AjusteEndosso, "
    lInsert = lInsert & " Alteroufranquia, "
    lInsert = lInsert & " VAL_PR_Orig, "
    lInsert = lInsert & " PremioInformado, "
    lInsert = lInsert & " TextoFranquia"

    'ProjetoFormularios
    lInsert = lInsert & " , Compl_Desc_Cobertura"
    'Regra Batente
    lInsert = lInsert & " , Val_Pr_Batente"
    'Mult Franquia
    lInsert = lInsert & " , Per_Desc_Franq"
    
    ''2017-02-24: Nao incluir se for FTP ou LOG
    If UCase(pNomeTabImport) <> "FTP_" And UCase(pNomeTabImport) <> "LOG_" Then
        lInsert = lInsert & ", Val_IS_ORIG, Per_Franq_ORIG "
    End If
    
    If UCase(pNomeTabImport) <> "FTP_" And UCase(pNomeTabImport) <> "LOG_" Then
        lInsert = lInsert & ", Per_Desc_Coml_Web "
    End If
    
    If UCase(pNomeTabImport) <> "FTP_" And UCase(pNomeTabImport) <> "LOG_" Then
        lInsert = lInsert & ", Tip_Origem_C4 "
    End If
    
    If UCase(pNomeTabImport) <> "FTP_" And UCase(pNomeTabImport) <> "LOG_" Then
        lInsert = lInsert & ", RA_C4 "
    End If
    
    If UCase(pNomeTabImport) <> "FTP_" And UCase(pNomeTabImport) <> "LOG_" Then
        lInsert = lInsert & ", VAL_IS_RA "
    End If

    ''2016-10-19~2018-02-28: VerbaCom
    If UCase(pPropTabela) Like "*SYASPORTAL*" Or UCase(pPropTabela) Like "*RAMOSDIVERSOS*" Then
        lInsert = lInsert & " , Val_VerbaComProp"
    End If

    lInsert = lInsert & ") VALUES ("
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero) & ","
    If pNomeTabImport = "FTP_" Or UCase(pNomeTabImport) = "LOG_" Then
        lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade) & ","
    End If
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Cobert) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Ped) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Org) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Ramo) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Apol) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Endo) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Tab_Cobert) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Num_Meses_PI) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Formula) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.COD_CLASSE) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Val_Fator_Franq) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Alterar_Cobert) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Val_Risco_Cob) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.NUM_PI) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.AjusteEndosso) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.AlterouFranquia) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Val_Pr_Orig) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.PremioInformado) & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.TextoFranquia)

    'ProjetoFormularios
    lInsert = lInsert & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Compl_Desc_Cobertura)
    'Regra Batente
    lInsert = lInsert & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.VAL_PR_BATENTE)
    'Multi franquia
    lInsert = lInsert & ","
    lInsert = lInsert & M46V999.gfFormatarCamposSQL(pobjstc.Per_Desc_Franq)

    ''2017-02-24: Nao incluir se for FTP ou LOG.
    ''Se DLL aceitar o valor enviado, se KIT copiar valor original
    If UCase(pNomeTabImport) <> "FTP_" And UCase(pNomeTabImport) <> "LOG_" Then
        If UCase(App.EXEName) Like "D46O003*" Then
            lInsert = lInsert & "," & M46V999.gfFormatarCamposSQL(pobjstc.Val_IS_Orig)
            lInsert = lInsert & "," & M46V999.gfFormatarCamposSQL(pobjstc.Per_Franq_Orig)
        Else
            lInsert = lInsert & "," & M46V999.gfFormatarCamposSQL(pobjstc.Val_IS)
            lInsert = lInsert & "," & M46V999.gfFormatarCamposSQL(pobjstc.Per_Franq)
        End If
    End If
    
    If UCase(pNomeTabImport) <> "FTP_" And UCase(pNomeTabImport) <> "LOG_" Then
        lInsert = lInsert & "," & M46V999.gfFormatarCamposSQL(pobjstc.Per_Desc_Coml_Web)
    End If
    
    If UCase(pNomeTabImport) <> "FTP_" And UCase(pNomeTabImport) <> "LOG_" Then
        lInsert = lInsert & "," & M46V999.gfFormatarCamposSQL(pobjstc.Tip_Origem_C4)
    End If
    
    If UCase(pNomeTabImport) <> "FTP_" And UCase(pNomeTabImport) <> "LOG_" Then
        lInsert = lInsert & "," & M46V999.gfFormatarCamposSQL(pobjstc.RA_C4)
    End If
    
    If UCase(pNomeTabImport) <> "FTP_" And UCase(pNomeTabImport) <> "LOG_" Then
        lInsert = lInsert & "," & M46V999.gfFormatarCamposSQL(pobjstc.VAL_IS_RA)
    End If

    ''2016-10-19~2018-02-28: VerbaCom
    If UCase(pPropTabela) Like "*SYASPORTAL*" Or UCase(pPropTabela) Like "*RAMOSDIVERSOS*" Then
        lInsert = lInsert & ", " & M46V999.gfFormatarCamposSQL(pobjstc.Val_VerbaComProp)
    End If
    
    lInsert = lInsert & ")"

    mfFormatarInclusaoSQL = lInsert

End Function


Private Function mfFormatarExclusaoSQL(pobjstc As stcA46V704B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String, _
                                       Optional ByVal CodTab As Integer) As String
    Dim lDelete As String                'Query.

    If CodTab = 2 Then GoTo Tab_Ped_Cobert_Compl

    lDelete = "DELETE FROM " & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert WHERE "
    lDelete = lDelete & " Nosso_Numero= " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lDelete = M46V999.gfWhereAnd(lDelete) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If
    If pobjstc.Num_Item > 0 Then
        lDelete = lDelete & " AND Num_Item = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item)
    End If

    mfFormatarExclusaoSQL = lDelete
    Exit Function

Tab_Ped_Cobert_Compl:
    ''2016-10-19~2018-02-28: VerbaCom
    lDelete = ""
    If pobjstc.Val_VerbaComProp > 0 Then
        If UCase(pPropTabela) Like "*SYASPORTAL*" Or UCase(pPropTabela) Like "*RAMOSDIVERSOS*" Then
            lDelete = "insert into RamosDiversos..Tab_VerbaComMovimento(Dat_Mov, Cod_Corr, Nosso_Numero, Val_VerbaCom, Cod_Acao)"
            lDelete = lDelete & " values (" & Format$(Now, "yyyymmdd") & ", " & M46V999.gfFormatarCamposSQL(pobjstc.ObjPai.ObjPai.objstcCtrlEmis.Cod_Corr) & ", " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero) & ", " & M46V999.gfFormatarCamposSQL(pobjstc.Val_VerbaComProp) & ", 80); "
            lDelete = lDelete & Chr(13) & Chr(10)
        End If
    End If

    lDelete = "DELETE FROM " & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert_Compl  WHERE "
    lDelete = lDelete & " Nosso_Numero= " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lDelete = M46V999.gfWhereAnd(lDelete) & " Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If
    If pobjstc.Num_Item > 0 Then
        lDelete = lDelete & " AND Num_Item = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item)
    End If

    mfFormatarExclusaoSQL = lDelete
End Function

Private Function mfFormatarConsultaSQL(ByRef pobjstc As stcA46V704B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As String
    Dim lSelect As String                'Query.

    lSelect = "SELECT " & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert.*, "
    lSelect = lSelect & " " & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert_Compl.Num_Meses_PI, "
    lSelect = lSelect & " " & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert_Compl.Val_Fator_Franq, "
    lSelect = lSelect & " " & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert_Compl.Formula, "
    lSelect = lSelect & " " & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert_Compl.Cod_Classe, "
    lSelect = lSelect & " " & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert_Compl.Alterar_Cobert, "
    lSelect = lSelect & " " & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert_Compl.Val_Risco_Cob, "
    'MarcoFranquia
    lSelect = lSelect & " " & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert_Compl.AjusteEndosso, "
    lSelect = lSelect & " " & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert_Compl.Alteroufranquia, "
    lSelect = lSelect & " " & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert_Compl.PremioInformado, "
    lSelect = lSelect & " " & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert_Compl.Val_PR_Orig, "
    lSelect = lSelect & " " & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert_Compl.TextoFranquia, "
    lSelect = lSelect & " " & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert_Compl.PremioInformado "

    'ProjetoFormularios
    lSelect = lSelect & ","
    lSelect = lSelect & " " & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert_Compl.Compl_Desc_Cobertura,"
    lSelect = lSelect & " " & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert_Compl.Val_Pr_Batente,"
    
    'Multiplo Franquia
    lSelect = lSelect & " " & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert_Compl.Per_Desc_Franq"
    
    ''2017-02-24
    If UCase(pNomeTabImport) <> "FTP_" And UCase(pNomeTabImport) <> "LOG_" Then
        lSelect = lSelect & ", " & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert_Compl.Val_IS_Orig, Tab_Ped_Cobert_Compl.Per_Franq_Orig "
    End If
    
    If UCase(pNomeTabImport) <> "FTP_" And UCase(pNomeTabImport) <> "LOG_" Then
        lSelect = lSelect & ", " & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert_Compl.Per_Desc_Coml_Web "
    End If
    
    If UCase(pNomeTabImport) <> "FTP_" And UCase(pNomeTabImport) <> "LOG_" Then
        lSelect = lSelect & ", " & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert_Compl.Tip_Origem_C4 "
    End If
    
    If UCase(pNomeTabImport) <> "FTP_" And UCase(pNomeTabImport) <> "LOG_" Then
        lSelect = lSelect & ", " & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert_Compl.RA_C4 "
    End If
    
    If UCase(pNomeTabImport) <> "FTP_" And UCase(pNomeTabImport) <> "LOG_" Then
        lSelect = lSelect & ", " & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert_Compl.VAL_IS_RA "
    End If

    ''2016-10-19~2018-02-28: VerbaCom
    If UCase(pPropTabela) Like "*SYASPORTAL*" Or UCase(pPropTabela) Like "*RAMOSDIVERSOS*" Then
        lSelect = lSelect & ", " & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert_Compl.Val_VerbaComProp "
    End If

    lSelect = lSelect & " From "
    lSelect = lSelect & "  " & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert with(nolock) LEFT JOIN     " & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert_Compl  with(nolock) ON"
    lSelect = lSelect & " (" & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert.COD_COBERT   = " & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert_Compl.COD_COBERT) AND"
    lSelect = lSelect & " (" & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert.NUM_ITEM     = " & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert_Compl.NUM_ITEM) AND"
    If M46V999.gfPreenchido(pobjstc.nosso_numero) Then
        lSelect = lSelect & " (" & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert.NOSSO_NUMERO = " & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert_Compl.NOSSO_NUMERO)"
    Else
        lSelect = lSelect & " (" & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert.Num_PI = " & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert_Compl.Num_PI)"
    End If

    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lSelect = lSelect & " AND (" & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert.Cod_Unicidade = " & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert_Compl.Cod_Unicidade)"
    End If
    If pobjstc.Cod_Cobert > 0 Then
        lSelect = lSelect & " Where  " & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert.Cod_Cobert = " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Cobert)
    End If
    If M46V999.gfPreenchido(pobjstc.nosso_numero) Then
        lSelect = M46V999.gfWhereAnd(lSelect) & " " & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert.Nosso_Numero= " & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    End If
    If pNomeTabImport = "FTP_" And pobjstc.Cod_Unicidade <> 0 Then
        lSelect = M46V999.gfWhereAnd(lSelect) & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert.Cod_Unicidade= " & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unicidade)
    End If
    If pobjstc.Num_Item > 0 Then
        lSelect = M46V999.gfWhereAnd(lSelect) & " " & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert.Num_Item = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item)
    End If
    If pobjstc.Num_Apol <> 0 Then
        lSelect = M46V999.gfWhereAnd(lSelect) & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert.Num_Apol = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Apol)
        lSelect = M46V999.gfWhereAnd(lSelect) & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert.num_endo = " & M46V999.gfFormatarCamposSQL(pobjstc.Num_Endo)
    End If
    lSelect = lSelect & " order by " & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert.num_item, " & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert.tab_cobert, " & pPropTabela & pNomeTabImport & "Tab_Ped_Cobert.cod_cobert"
    mfFormatarConsultaSQL = lSelect
End Function
Private Sub mpLimparColecao()

    Set Colecao = Nothing
    Set Colecao = New Collection
    Set ColecaoNovaFranquia = Nothing
    Set ColecaoNovaFranquia = New Collection
End Sub
Private Sub mpCarregarAtributos(ByRef pobjstc As stcA46V704B, prst As clsYasRecordSet, Optional ByVal pParcial As Boolean = False)

    If M46V999.gCodProduto >= 1080 And Not pobjstc.ObjPai.ObjPai Is Nothing Then
        If pobjstc.ObjPai.ObjPai.Cod_Estip < 2000 And pobjstc.ObjPai.ObjPai.Cod_Produto >= 1080 Then 'NÃO FAZER DE-> PARA DE PLANINHO
                pobjstc.Cod_Cobert = M46V999.gfBuscarDePara(M46V999.gfTratarNulo(prst("Cod_Cobert"), tpoNumerico), 1)
        Else
                pobjstc.Cod_Cobert = M46V999.gfTratarNulo(prst("Cod_Cobert"), tpoNumerico)
        End If
    Else
        pobjstc.Cod_Cobert = M46V999.gfTratarNulo(prst("Cod_Cobert"), tpoNumerico)
    End If
    pobjstc.nosso_numero = M46V999.gfTratarNulo(prst("NOsso_Numero"), tpoString)
    pobjstc.Num_Ped = M46V999.gfTratarNulo(prst("Num_Ped"), tpoNumerico)
    pobjstc.Cod_Org = M46V999.gfTratarNulo(prst("Cod_Org"), tpoNumerico)
    pobjstc.Cod_Ramo = M46V999.gfTratarNulo(prst("Cod_Ramo"), tpoNumerico)
    pobjstc.Num_Apol = M46V999.gfTratarNulo(prst("Num_Apol"), tpoNumerico)
    pobjstc.Num_Endo = M46V999.gfTratarNulo(prst("Num_Endo"), tpoNumerico)
    pobjstc.Num_Item = M46V999.gfTratarNulo(prst("Num_item"), tpoNumerico)
    pobjstc.Tab_Cobert = M46V999.gfTratarNulo(prst("Tab_Cobert"), tpoNumerico)
    pobjstc.Val_IS = M46V999.gfTratarNulo(prst("Val_IS"), tpoNumerico)
    pobjstc.Per_Taxa = M46V999.gfTratarNulo(prst("Per_Taxa"), tpoNumerico)
    pobjstc.Per_Desc_Coml = M46V999.gfTratarNulo(prst("Per_Desc_Coml"), tpoNumerico)
    pobjstc.Per_Desc_Exp = M46V999.gfTratarNulo(prst("Per_Desc_Exp"), tpoNumerico)
    pobjstc.VAL_PR = M46V999.gfTratarNulo(prst("Val_PR"), tpoNumerico)
    pobjstc.Val_Pr_Anu = M46V999.gfTratarNulo(prst("Val_Pr_Anu"), tpoNumerico)
    pobjstc.Val_Pr_Tarif = M46V999.gfTratarNulo(prst("Val_Pr_Tarif"), tpoNumerico)
    pobjstc.Val_Pr_Tarif_Anu = M46V999.gfTratarNulo(prst("Val_Pr_Tarif_Anu"), tpoNumerico)
    pobjstc.Val_Pr_Resg = M46V999.gfTratarNulo(prst("Val_Pr_Resg"), tpoNumerico)
    pobjstc.Val_Pr_Resg_Anu = M46V999.gfTratarNulo(prst("Val_Pr_Resg_Anu"), tpoNumerico)
    pobjstc.Per_Franq = M46V999.gfTratarNulo(prst("Per_Franq"), tpoNumerico)
    pobjstc.Val_Franq_Min = M46V999.gfTratarNulo(prst("Val_Franq_Min"), tpoNumerico)
    pobjstc.Val_Franq_Max = M46V999.gfTratarNulo(prst("Val_Franq_Max"), tpoNumerico)
    pobjstc.Num_Meses_PI = M46V999.gfTratarNulo(prst("Num_Meses_PI"), tpoNumerico)
    pobjstc.COD_CLASSE = M46V999.gfTratarNulo(prst("Cod_Classe"), tpoNumerico)
    pobjstc.Formula = M46V999.gfTratarNulo(prst("Formula"), tpoString)
    pobjstc.Val_Fator_Franq = M46V999.gfTratarNulo(prst("Val_Fator_Franq"), TpoValor)
    pobjstc.Alterar_Cobert = M46V999.gfTratarNulo(prst("Alterar_Cobert"), tpoString)
    pobjstc.Val_Risco_Cob = M46V999.gfTratarNulo(prst("Val_Risco_Cob"), TpoValor)
    pobjstc.NUM_PI = M46V999.gfTratarNulo(prst("Num_PI"), tpoNumerico)
    pobjstc.AjusteEndosso = M46V999.gfTratarNulo(prst("AjusteEndosso"), tpoString)
    pobjstc.AlterouFranquia = M46V999.gfTratarNulo(prst("AlterouFranquia"), tpoString)
    pobjstc.TextoFranquia = M46V999.gfTratarNulo(prst("TextoFranquia"), tpoString)
    pobjstc.Val_Pr_Orig = M46V999.gfTratarNulo(prst("Val_PR_Orig"), TpoValor)
    
    ''2017-02-24: Novos campos de Val_Is_Orig e Per_Franq_Orig
    pobjstc.Val_IS_Orig = M46V999.gfTratarNulo(prst("Val_IS_ORIG"), TpoValor)
    pobjstc.Per_Franq_Orig = M46V999.gfTratarNulo(prst("Per_Franq_ORIG"), TpoValor)

    If Trim(prst("PremioInformado")) = "" Then
        pobjstc.PremioInformado = False
    Else
        pobjstc.PremioInformado = M46V999.gfTratarNulo(prst("PremioInformado"), TpoBoolean)
    End If
    'ProjetoFormularios
    On Error Resume Next
    pobjstc.Compl_Desc_Cobertura = M46V999.gfTratarNulo(prst("Compl_Desc_Cobertura"), tpoString)
    'Regra Batente
    pobjstc.VAL_PR_BATENTE = M46V999.gfTratarNulo(prst("Val_Pr_Batente"), TpoValor)
    'Mult Franquia
    pobjstc.Per_Desc_Franq = M46V999.gfTratarNulo(prst("Per_Desc_Franq"), TpoValor)
    pobjstc.Per_Desc_Coml_Web = M46V999.gfTratarNulo(prst("Per_Desc_Coml_Web"), tpoNumerico)
    pobjstc.Tip_Origem_C4 = M46V999.gfTratarNulo(prst("Tip_Origem_C4"), tpoNumerico)
    pobjstc.RA_C4 = M46V999.gfTratarNulo(prst("RA_C4"), tpoNumerico)
    pobjstc.VAL_IS_RA = M46V999.gfTratarNulo(prst("VAL_IS_RA"), tpoNumerico)
    ''2016-10-19~2018-02-28: VerbaCom
    pobjstc.Val_VerbaComProp = M46V999.gfTratarNulo(prst("Val_VerbaComProp"), TpoValor)
End Sub

Public Function mfIncluir(ByRef pobjstc As stcA46V704B, _
                          ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String, _
                          ByVal NaoIncluirCompl As Boolean, _
                          ByVal NaoIncluirCober As Boolean, tagDynaTrace As String) As Boolean
                          
    ' Valores padrão
    ' NaoIncluirCompl = false
    ' NaoIncluirCober = false

    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    Dim i As Integer
    methodName = "mfIncluir"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)
    
    Dim lsSQl As String           'Query.

    On Local Error GoTo ControleErro

    If Not mfConectar Then GoTo ControleErro

    'Caso coberturas com de->para de n para 1 pegar a de maior IS
    For i = 1 To Colecao.Count
        If i > Colecao.Count Then
            Exit For
        End If
        
        If Colecao(i).Class_ID = pobjstc.Class_ID Then
            'CASO A COBERTURA COM MENOR IS JA FOI INCLUIDA, EXCLUIR PARA ADICIONAR A COM MAIOR IS.
            If Colecao(i).Val_IS <= pobjstc.Val_IS Then
                Colecao.Remove (pobjstc.Class_ID)
                lsSQl = "DELETE FROM " & pPropTabela & "TAB_PED_COBERT WHERE NOSSO_NUMERO = '" & pobjstc.nosso_numero & "'"
                lsSQl = lsSQl & " AND COD_COBERT = " & pobjstc.Cod_Cobert & " AND NUM_ITEM = " & pobjstc.Num_Item
                
                If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then GoTo ControleErro
        
                lsSQl = "DELETE FROM " & pPropTabela & "TAB_PED_COBERT_COMPL WHERE NOSSO_NUMERO = '" & pobjstc.nosso_numero & "'"
                lsSQl = lsSQl & " AND COD_COBERT = " & pobjstc.Cod_Cobert & " AND NUM_ITEM = " & pobjstc.Num_Item
                
                If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then GoTo ControleErro
            Else
                GoTo Proximo
            End If
        End If
        
        'EXCLUIR COBERTURAS QUE NÃO EXISTEM MAIS NO NOVO PRODUTO
        If pobjstc.Cod_Cobert < 1000 And pobjstc.Cod_Cobert <> 1 And M46V999.gNum_Apolice > 0 And M46V999.gCodProduto >= 1080 And M46V999.M46V111.gCod_plano < 2000 Then
             'pobjstc.ColCobertExcluidas.Add Item:=CStr(pobjstc.Cod_Cobert) & "-" & CStr(pobjstc.gNomeCobertura) & ";"
            GoTo Proximo
        End If
    Next

    If NaoIncluirCober <> True Then
        lsSQl = mfFormatarInclusaoSQL(pobjstc, pPropTabela, pNomeTabImport)
        If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then GoTo ControleErro
    End If

    If NaoIncluirCompl <> True Then
        lsSQl = mfFormatarInclusaoSQL(pobjstc, pPropTabela, pNomeTabImport, 2)
        If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then GoTo ControleErro
    End If

    Colecao.Add Item:=pobjstc, Key:=pobjstc.Class_ID
    
Proximo:
    mfIncluir = True
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
Exit Function
ControleErro:
    mfIncluir = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V704A.mfIncluir ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V704A.mfIncluir ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, UltimoErro, 2, Class_Name, methodName)
End Function
Public Function mfExcluir(ByRef pobjstc As stcA46V704B, _
                          ByVal pPropTabela As String, _
                          ByVal pNomeTabImport As String, _
                          Optional NaoExcluirCompl As Boolean, _
                          Optional Cod_Cobert As Integer = 0) As Boolean
    Dim lsSQl As String           'Query.

    On Error GoTo ControleErro

    If Not mfConectar Then GoTo ControleErro
    
    If Cod_Cobert = 0 Then
        If Not mfExcluirDependentes(pobjstc, pPropTabela, pNomeTabImport) Then GoTo ControleErro
    End If

    lsSQl = mfFormatarExclusaoSQL(pobjstc, pPropTabela, pNomeTabImport)
    If Cod_Cobert > 0 Then
        lsSQl = lsSQl & " AND Cod_Cobert = " & M46V999.gfFormatarCamposSQL(Cod_Cobert)
    End If
    If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then GoTo ControleErro

    If NaoExcluirCompl <> True Then
        lsSQl = mfFormatarExclusaoSQL(pobjstc, pPropTabela, pNomeTabImport, 2)
        If Cod_Cobert > 0 Then
            lsSQl = lsSQl & " AND Cod_Cobert = " & M46V999.gfFormatarCamposSQL(Cod_Cobert)
        End If
        If M46V999.gfExeSQL(ConexaoBD, lsSQl, UltimoErro) <> 0 Then GoTo ControleErro
    End If

    If Colecao.Count > 0 And Cod_Cobert = 0 Then
        Colecao.Remove (pobjstc.Class_ID)
        Set pobjstc = Nothing
    End If
    mfExcluir = True
    Exit Function
ControleErro:
    mfExcluir = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V704A.mfExcluir ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V704A.mfExcluir ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
End Function
Public Function mfCarregarColecao(ByRef pobjstc As stcA46V704B, _
                                  ByVal pPropTabela As String, _
                                  ByVal pNomeTabImport As String) As Boolean
    Dim lobjstc As stcA46V704B          'Objeto da classe de estrutura stcA46V704B.
    Dim lrst As clsYasRecordSet   'Registro.
    Dim lsSQl As String             'SQL.
    Dim i As Integer

    On Error GoTo ControleErro

    If Not mfConectar Then GoTo ControleErro

    If M46V999.gfObtRegistro(ConexaoBD, mfFormatarConsultaSQL(pobjstc, pPropTabela, pNomeTabImport), _
                             lrst, UltimoErro) = False Then GoTo ControleErro

    Call mpLimparColecao
    Do While Not lrst.EOF
        Set lobjstc = New stcA46V704B
        Set lobjstc.M46V999 = M46V999

        Set lobjstc.ObjPai = pobjstc.ObjPai

        Call mpCarregarAtributos(lobjstc, lrst)
        If Not mfCarregarDependentes(lobjstc, pPropTabela, pNomeTabImport) Then GoTo ControleErro
        
        For i = 1 To Colecao.Count
            If i > Colecao.Count Then
                Exit For
            End If
            
            If Colecao(i).Class_ID = lobjstc.Class_ID Then
                If Colecao(i).Val_IS <= lobjstc.Val_IS Then
                    Colecao.Remove (lobjstc.Class_ID)
                Else
                GoTo Proximo
                End If
            End If
        Next
            
            'EXCLUIR COBERTURAS QUE NÃO EXISTEM MAIS NO NOVO PRODUTO
            If Not lobjstc.ObjPai.ObjPai Is Nothing Then
                'RENOVAÇÕES YASUDA
                If lobjstc.Cod_Cobert < 1000 And lobjstc.Cod_Cobert <> 1 And lobjstc.ObjPai.ObjPai.Num_Apol_An > 0 And lobjstc.ObjPai.ObjPai.Cod_Produto >= 1080 And lobjstc.ObjPai.ObjPai.Cod_Estip < 2000 Then
                    GoTo Proximo
                End If
            
                'RENOVAÇÕES MARITIMA
                If lobjstc.ObjPai.ObjPai.objstcCtrlEmis.Num_Apol_Mar > 0 Then
                    'seto as variaveis globais para carregar os dados adicionais
                    M46V999.gCod_Ramo = lobjstc.Cod_Ramo
                    M46V999.M46V111.gCod_plano = lobjstc.Tab_Cobert
                    M46V999.gCodProduto = lobjstc.ObjPai.ObjPai.Cod_Produto
                    
                    Call M46V999.BuscaDadosAdicionais(ConexaoBD, UltimoErro)
                    
                    If M46V999.gRst_Tab_Cobert.State = adStateOpen Then
                        M46V999.gRst_Tab_Cobert.Filter = adFilterNone
    
                        M46V999.gRst_Tab_Cobert.Filter = "COD_Ramo = " & lobjstc.Cod_Ramo & " AND COD_TARIF_INI = " & lobjstc.ObjPai.ObjPai.Cod_Produto & " AND tab_cobert = " & lobjstc.Tab_Cobert & " AND cod_cobert = " & lobjstc.Cod_Cobert
                
                        If M46V999.gRst_Tab_Cobert.EOF Then
                            GoTo Proximo
                        End If
                    End If
                End If
            End If
                
        Colecao.Add Item:=lobjstc, Key:=lobjstc.Class_ID
Proximo:
        lrst.MoveNext
    Loop
    mfCarregarColecao = True

SAIDA:
    GoSub FechaObj
    Exit Function

ControleErro:
    mfCarregarColecao = False
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V704A.mfCarregarColecao ERRO - " & Err & " " & Err.Description & _
                                 vbCrLf & " SQL: " & lsSQl)
    Else
        Call M46V999.gpGraLog(0, "A46V704A.mfCarregarColecao ERRO - " & UltimoErro & _
                                 vbCrLf & " SQL: " & lsSQl)
    End If
    GoTo SAIDA
FechaObj:
    M46V999.gRst_Tab_Cobert.Filter = adFilterNone
    Set lobjstc = Nothing
    Set lrst = Nothing
    Return
End Function

Private Sub Class_Initialize()

    Set Colecao = New Collection
    Set mClsM46V999 = Nothing
End Sub
Private Sub Class_Terminate()
    Set Colecao = Nothing
    Set mClsM46V999 = Nothing
End Sub
Private Function mfConectar() As Boolean
    mfConectar = False
    If ConexaoBD Is Nothing Then
        If M46V999.gfAbrBasDados("P00MPED", M46V999.gTipP00MPED, M46V999.gEndP00MPED, ConexaoBD, _
                                 UltimoErro) = False Then Exit Function
    ElseIf ConexaoBD.State <> 1 Then
        If M46V999.gfAbrBasDados("P00MPED", M46V999.gTipP00MPED, M46V999.gEndP00MPED, ConexaoBD, _
                                 UltimoErro) = False Then Exit Function
    End If
    mfConectar = True
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


Private Function mfCarregarDependentes(ByRef pobjstc As stcA46V704B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As Boolean
    Dim lobjMultFraq As clsA46V738A    'Objeto da classe clsA46V738A (TAB_PED_COBERT_FRANQUIA)
    Dim lobjstcMultFraq As stcA46V738B    'Objeto da classe de estrutura stcA46V738B (TAB_PED_COBERT_FRANQUIA)

    On Error GoTo ControleErro
    mfCarregarDependentes = False

    'ProjetoMultiploFranquias
    'Carregar coleção de nova franquia.

    Set lobjMultFraq = New clsA46V738A
    Set lobjMultFraq.M46V999 = M46V999

    Set lobjstcMultFraq = New stcA46V738B
    Set lobjstcMultFraq.M46V999 = M46V999

    lobjstcMultFraq.nosso_numero = pobjstc.nosso_numero
    lobjstcMultFraq.NUM_PI = pobjstc.NUM_PI
    lobjstcMultFraq.Num_Item = pobjstc.Num_Item
    lobjstcMultFraq.Num_Apol = pobjstc.Num_Apol
    lobjstcMultFraq.Num_Endosso = pobjstc.Num_Endo
    lobjstcMultFraq.Cod_Cobert = pobjstc.Cod_Cobert
    
    Set lobjstcMultFraq.ObjPai = pobjstc

    Set lobjMultFraq.ConexaoBD = Me.ConexaoBD
    If UCase(pPropTabela) <> "GED.DBO." Then
        If Not lobjMultFraq.mfCarregarColecao(lobjstcMultFraq, pPropTabela, pNomeTabImport) Then
            UltimoErro = lobjMultFraq.UltimoErro
            GoTo ControleErro
        End If
    End If
    Set Me.ConexaoBD = lobjMultFraq.ConexaoBD
    Set pobjstc.ColecaoNovaFranquia = lobjMultFraq.Colecao

    pobjstc.EhGrandesRiscos = False
    If Not pobjstc.ObjPai.ObjPai Is Nothing Then
        If pobjstc.Cod_Cobert = 1 Then
            If M46V999.M46V103.mpVerificaFranqGrdRisco(pobjstc.Cod_Ramo, pobjstc.ObjPai.ObjPai.Cod_Estip, pobjstc.ObjPai.ObjPai.Cod_Produto, pobjstc.Val_IS) Then
                pobjstc.EhGrandesRiscos = True
            End If
        End If
    End If

    mfCarregarDependentes = True

SAIDA:
    GoSub FechaObj
    Exit Function
ControleErro:
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V704A.mfCarregarDependentes ERRO - " & Err & " " & Err.Description)
    Else
        Call M46V999.gpGraLog(0, "A46V704A.mfCarregarDependentes ERRO - " & UltimoErro)
    End If
    GoTo SAIDA
FechaObj:
    Set lobjMultFraq = Nothing
    Set lobjstcMultFraq = Nothing

    Return
End Function

Private Function mfExcluirDependentes(ByRef pobjstc As stcA46V704B, _
                                       ByVal pPropTabela As String, _
                                       ByVal pNomeTabImport As String) As Boolean
    Dim lobjMultFraq As clsA46V738A    'Objeto da classe clsA46V738A (TAB_PED_COBERT_FRANQUIA)
    Dim lobjstcMultFraq As stcA46V738B    'Objeto da classe de estrutura stcA46V738B (TAB_PED_COBERT_FRANQUIA)

    On Error GoTo ControleErro
    mfExcluirDependentes = False

    'ProjetoMultiploFranquias
    'Exclui coleção de nova franquia.

    Set lobjMultFraq = New clsA46V738A
    Set lobjMultFraq.M46V999 = M46V999

    Set lobjstcMultFraq = New stcA46V738B
    Set lobjstcMultFraq.M46V999 = M46V999

    lobjstcMultFraq.nosso_numero = pobjstc.nosso_numero
    lobjstcMultFraq.Num_Item = pobjstc.Num_Item
    lobjstcMultFraq.Num_Apol = pobjstc.Num_Apol
    lobjstcMultFraq.Num_Endosso = pobjstc.Num_Endo
    
    Set lobjstcMultFraq.ObjPai = pobjstc

    Set lobjMultFraq.ConexaoBD = Me.ConexaoBD
    If UCase(pPropTabela) <> "GED.DBO." Then
        If Not lobjMultFraq.mfExcluir(lobjstcMultFraq, pPropTabela, pNomeTabImport) Then
            UltimoErro = lobjMultFraq.UltimoErro
            GoTo ControleErro
        End If
    End If
    Set Me.ConexaoBD = lobjMultFraq.ConexaoBD
    Set pobjstc.ColecaoNovaFranquia = lobjMultFraq.Colecao

    mfExcluirDependentes = True

SAIDA:
    GoSub FechaObj
    Exit Function
ControleErro:
    If M46V999.gfPreenchido(UltimoErro) = False Then
        UltimoErro = Err.Description
        Call M46V999.gpGraLog(0, "A46V704A.mfExcluirDependentes ERRO - " & Err & " " & Err.Description)
    Else
        Call M46V999.gpGraLog(0, "A46V704A.mfExcluirDependentes ERRO - " & UltimoErro)
    End If
    GoTo SAIDA
FechaObj:
    Set lobjMultFraq = Nothing
    Set lobjstcMultFraq = Nothing

    Return
End Function
