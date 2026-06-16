VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6405
   ClientLeft      =   2535
   ClientTop       =   1860
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6405
   ScaleWidth      =   6795
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   1095
      Left            =   1560
      TabIndex        =   1
      Top             =   2640
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1095
      Left            =   1560
      TabIndex        =   0
      Top             =   1080
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load_Novo()
    Dim a As New clsA46O003A
    Dim gM46V999 As New clsM46V999
    Dim lBanco As ADODB.Connection
    Dim lMensagem As String
    Dim lRs As clsYasRecordSet
    Dim lSelect As String
    Dim normas  As New Collection
    Dim alertas As New Collection
    Dim erros   As New Collection
    Dim lTabPed  As stcA46V702B
    Dim lRetornoEhErro As Boolean
    Dim objstc As stcA46V702B
    Dim pMensagensRetorno As String
    Dim t99 As New clsM46V999
     
    a.gfAutenticarUsuario "test28", ""
    'a.gfAutenticarUsuario "ysnt47", "dantiele12"
    'a.gfAutenticarUsuario "ysnt13", "yasuda01"
    
    gM46V999.gfObtSYASMULTINI ("")
    Set gM46V999.DynaTrace = New clsMDynaTrace
    'a.objstcUsuario.nom_user
    
    t99.gfVersao
    
    pMensagensRetorno = a.gfTransmitirProposta("01721252403121410550", False, "")
    
    Set objstc = a.gfProcuraNossoNumero("01717233817263576697", False, "SYASPORTAL", "", 0)
    
    
    'Set lTabPed = a.gfCalcular("01714359388832454867", normas, alertas, erros, False, "")
    
    If gM46V999.gfAbrBasDados("SYASPORTAL", gM46V999.gTipSyasPortal, gM46V999.gEndSyasPortal, lBanco, lMensagem) = False Then
        MsgBox lMensagem, vbCritical
        GoTo Saida
    End If
    
    Set gM46V999.M46V111 = New clsM46V111
    Set objstc.M46V999 = gM46V999
    
    If Not gM46V999.gfTransmissaoInterna(1, objstc, lMensagem, lBanco, lRetornoEhErro) Then

    End If
    
    
    lSelect = "SELECT Nosso_Numero FROM syasportal..tab_ped WHERE Nosso_Numero = '01717158791036089250'"
    
    If Not gM46V999.gfObtRegistro(lBanco, lSelect, lRs, lMensagem) Then
        MsgBox lMensagem, vbCritical
    End If
    
    Do While Not lRs.EOF
        Set lTabPed = a.gfCalcular("01717956237523987950", normas, alertas, erros, False, "")
        lMensagem = a.gfTransmitirProposta("01717956237523987950", False, "")
        'lMensagem = a.gfTransmitirProposta(lRs!Nosso_Numero, False, "")
        lRs.MoveNext
    Loop

    MsgBox "Terminou"

Saida:
    
End Sub

Private Sub Form_Load()
'    Dim teste2 As String
'    Dim teste As clsA46O003A
'    Set teste = New clsA46O003A
'
'    teste2 = teste.gfProtocoloTransmissao("01516250292206630899", "", "")
    Dim a As New clsA46O003A
    Dim objstc  As stcA46V702B
    Dim U   As New stcA46V725B
    Dim pMensagensRetorno As String
    Dim normas  As New Collection
    Dim alertas As New Collection
    Dim erros   As New Collection
    Dim classe4 As Boolean
    Dim norma1 As New stcA46V708N
    Dim norma2 As New stcA46V708N
    Dim rel As stcA46V702rpt
    
    Set a = New clsA46O003A
    'a.gfAutenticarUsuario "ysnt13", "yasuda01"
    'a.gfAutenticarUsuario "ysnt47", "dantiele12"
    a.gfAutenticarUsuario "test28", ""
    Set U = a.objstcUsuario
    
    pMensagensRetorno = a.gfTransmitirProposta("01717233817263576697", False, "")
    pMensagensRetorno = a.UltimoErro

    'a.M46V999.gOrigemChamada = e_Origem_Kit
    Set objstc = a.gfProcuraNossoNumero("01709561273906882887", False, "SYASPORTAL", "", 0)
    'Set objstc = a.gfProcuraNossoNumero("01702460117971645535", True, "RAMOSDIVERSOS", "", 0)
    Set rel = objstc.getRpt
    
    Dim teste As New clsM46V998
    Dim parcela As Variant
    Dim lIOF As Single
    
    Set objstc.M46V999 = New clsM46V999
    Call teste.gfCalFracionamento(objstc, objstc.ColecaoPedLoc(1).Tip_Emissao, parcela, pMensagensRetorno, objstc.getProduto.ProJurIOF, objstc.Cod_Reg, lIOF, objstc.Val_Parc_Min)

    classe4 = a.M46V999.gfCopProposta("01706536144786663290", "51626659829881871520 ", "999999", pMensagensRetorno)
    
    Set objstc = a.gfSeguroNovo(112, 0, 0, "")
    
    
    pMensagensRetorno = a.gfTransmitirProposta("01713736088929454989", False, "")
    
    classe4 = a.gfExcluirItem("01626078408870614947", 2, "")
    
    norma1.COD_CLASSE_NORMA = 4
    norma1.COD_STAT_IRREG = 10
    norma1.Num_Item = 1
    norma1.COD_ERR_NORMA = "102.0.0.32"
    norma1.DSC_ERR_NORMA_EXT = "Alagamento e Inundacao - Solicitar cotação"
    norma1.DSC_ERR_NORMA_INT = norma1.DSC_ERR_NORMA_EXT
    normas.Add norma1
     
    norma2.COD_CLASSE_NORMA = 1
    norma2.COD_STAT_IRREG = 55
    norma2.Num_Item = 1
    norma2.COD_ERR_NORMA = "104.0.0.1112"
    norma2.DSC_ERR_NORMA_EXT = "Inspecao obrigatoria para cobertura 1112"
    norma2.DSC_ERR_NORMA_INT = norma1.DSC_ERR_NORMA_EXT
    normas.Add norma2
    
    Set normas = New Collection
    
    Set objstc = a.gfCalcular("01727754313827748590", normas, alertas, erros, False, "")
    
    Set objstc = a.gfProEndosso(1800166487, 1, False, 12, pMensagensRetorno, "")
    
    Set objstc = a.gfProEndosso(1811275028, 1, False, 12, pMensagensRetorno, "")


    
    a.M46V999.gfTransmissaoInterna 1, objstc, pMensagensRetorno
    
    
    'classe4 = a.gfSolicitarClasse4("01625852344206951324", "testes@sompo.com.br", "39764523", "teste", "teste da silva", pMensagensRetorno, "")
    
'    Dim rel As stcA46V702rpt
'    Set rel = objstc.getRpt
'
'    Dim reloc As stcA46V708rpt
'    Set reloc = objstc.ColecaoPedLoc(1).getRpt
'
'    pMensagensRetorno = reloc.ObservacoesNew
    
    
End Sub

Private Sub Command1_Click()
Dim a As New clsA46O003A
Dim U   As New stcA46V725B
Dim objstc  As stcA46V702B
Dim lRet    As String
Dim lColecaoA   As Collection
Dim lColecaoB   As Collection
Dim lColecaoC   As Collection
Dim lClasse4    As Boolean

 
 
 
'Dim lMensagem As String
'Dim parcela As Variant
'Dim mNossoNumero As String
'
'    Set a = New clsA46O003A
'    a.gfAutenticarUsuario ("ysnt15")
'    Set U = a.objstcUsuario
'    'Call a.gfSeguroNovo(113, 0, 1)
'    'Set objstc = a.gfProcuraNossoNumero("01320449544000016012")
'    'Call a.gfRetornaNovoItem(objstc, 3, 12)
'    Call a.gfProRenYasuda(11799025, 1, 0, lMensagem)
'    'Call a.gfProEndosso(11621012, 2, True, 1, lMensagem)
'
'    mNossoNumero = ""
'    Debug.Print mNossoNumero
'    Set lColecaoA = New Collection
'    Set lColecaoB = New Collection
'    Set lColecaoC = New Collection
'    mNossoNumero = "01320454529000016013"
'
'
'
'
'
'
''    Set objstc = a.gfProcuraNossoNumero("01300663507002631956")
''    Call a.gfRetornaNovoItem(objstc, 2, 4)
'    Set objstc = a.gfCalcular(mNossoNumero, lColecaoA, lColecaoB, lColecaoC, lClasse4)
'
'    Call a.gfTransmitirProposta(mNossoNumero, True)
'
'    Dim stcitem As stcA46V708B
'
'    Set stcitem = objstc.ColecaoPedLoc.Item(1)
'    Dim stccob As stcA46V704B
'    Set stccob = stcitem.ColecaoPedCobert.Item(10)
'
'Exit Sub
'
'    Set a = New clsA46O003A
'    a.gfAutenticarUsuario ("SYAS01")
'    Set U = a.objstcUsuario
'
'
''Set objstc = a.gfProcuraNossoNumero("01225040633000157018", True)
'lRet = a.gfProtocoloTransmissao("01225452541000016013")
'
'
'    Set lColecaoA = New Collection
'    Set lColecaoB = New Collection
'    Set lColecaoC = New Collection
'
'
'
'Exit Sub
'
'
'
'
'
''   objstc.Beneficiarios = ""
'
'
'    Set a = New clsA46O003A
'    a.gfAutenticarUsuario ("SYAS03")
'    Set U = a.objstcUsuario
'
''Set objstc = a.gfProcuraNossoNumero("01233159433000080018")
'    Set lColecaoA = New Collection
'    Set lColecaoB = New Collection
'    Set lColecaoC = New Collection
'    Call a.gfCalcular("01233159433000080018", lColecaoA, lColecaoB, lColecaoC, lClasse4)
'
'    Exit Sub
'
'    'Call a.gfImprimeObjeto(objstc)
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4377 | TipoUsuario=2/TipoAcesso=2/Corretor=1601/CodUserMsp=/prior_multi=0/syasportal.DBO.
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4378 |
' Dim lobjstcPedLoc   As stcA46V708B
' Dim lObjstcPedClau  As stcA46V703B
' Dim lobjstcPedCobert  As stcA46V704B
' Dim lobjstcPedCorr  As stcA46V705B
' Dim lobjstcPedCoss  As stcA46V706B
' Dim lobjstcPedEnder  As stcA46V707B
' Dim lobjstcPedJuros  As stcA46V711B
' Dim lobjstcPedCtrEmis As stcA46V712B
' Dim lobjstcPedParc As stcA46V713B
' Dim lobjstcPedObs As stcA46V714B
' Dim lobjstcPedBem As stcA46V715B
' Dim lobjstcPedOutroSeg As stcA46V716B
' Dim lobjstcPedInsp As stcA46V717B
' Dim lobjstcFator    As stcA46V720B
' Dim lobjstcPedEspec As stcA46V726B
' Dim lstcNorma       As stcA46V708N
' Dim lobjstcPed      As stcA46V702B
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4379 |
'
' Set lobjstcPed = New stcA46V702B
'    With lobjstcPed
' .NOSSO_NUMERO = "01133651010000016011"
' .Cod_Ramo = 113
' .Cod_Org = 0
' .Num_Ped = 2
' .Num_Apol = 0
' .Num_Endo = 0
' .Num_Apol_An = 11973275
' .Cod_Reg = 1
' .Cod_Bco_Cobr = 341
' .Cod_Agen_Angar = 0
' .Cod_Agen_Indicr = 0
' .Num_Cta_Indicr = 0
' .Cod_Unid_Prod = 251
' .Cod_Prod = 1692
' .Cod_Agen_Bas_Cor = 0
' .Cod_Dig_Agen_Bas_Cor = 0
' .Num_Cta_Bas_Cor = 0
' .Cod_Segurado = 7777023
' .Dig_Cod_Segurado = ""
' .Nom_Segurado = "ANDRE LUIZ LOPES DE OLIVEIRA (YASUDA)"
' .Dsc_Compl_Segurado = ""
' .Num_Cgc_Cpf = 28708974802#
' .Cod_Tip_Logr = ""
' .Nom_Logr = ""
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4380 |
' .Dsc_Compl_Logr = ""
' .Nom_Bairro = ""
' .Nom_Cid = ""
' .Sig_UF = ""
' .Num_CEP = 0
' .Num_Fone = 0
' .Cod_Estip = 1
' .Val_Coef_Cms_Pad = 0
' .Dat_Ini_Vig = 20111223
' .Dat_Ter_Vig = 20121223
' .Dat_Emis = 0
' .Cod_Moed = 0
' .Per_Juro_Frac = 0
' .Qtd_Parc = 10
' .Qtd_Anexo = 1
' .Cod_Iden_Praz = 1
' .Cod_Pr_Min_Endo = 1
' .Cod_Pr_Min_Parc = 1
' .Val_Pr_Antec = 0
' .Per_Adic_Form_Pag = 0
' .Per_Desc_Form_Pag = 0
' .Val_IS = 250000
' .Val_IS_2 = 0
' .VAL_PR = 207.14
' .Val_Pr_Ref_Tarif = 207.14
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4381 |
' .Val_Pr_Ref_Resg = 207.14
' .Dat_Ini_Viag = 0
' .Cod_Tip_Liber = 0
' .Cod_Iden_Alt_Corr = ""
' .Cod_User_Digit = ""
' .DAT_DIGIT = 20111202
' .HOR_DIGIT = 1416
' .Cod_User_Liber = "YSNT21"
' .Dat_Liber = 20111202
' .Hor_Liber = 1416
' .Cod_Sit_Ped = 2
' .Cod_Tip_Doc = 2
' .Cod_Dizer = 0
' .Qtd_Copia = 3
' .Val_Cust_Emis = 50
' .Val_Parc_Min = 0
' .Val_Juro = 0
' .Val_IOF = 18.9799995422363
' .Num_Prop = 0
' .Cod_Posi_Atual = ""
' .Dat_Mov_Ini = 0
' .Dat_Mov_Ter = 0
' .Tip_Cust_Emis = 1
' .Val_Cambio_Emis = 0
' .Iden = ""
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4382 |
' .NUM_PI = 1100000002
' .Cod_Produto = 1024
' .Cod_Protocolo = 0
' .Cod_Sequencia = 0
' .Tip_Pessoa = 2
' .Cod_Ramo_Ativ = 0
' .Cod_Doc = 1
' .Num_Doc = "30909907-9"
' .Org_Exp = ""
' .Dat_Emis_Doc = 20040510
' .Nom_Empresa = ""
' .Sexo = "M"
' .Dat_Pri_Calculo = 0
' .Ult_Tarifa = 0
' .Cod_Unicidade = 88260
' .Cod_Ver_Unici = 1
' .Val_PriParcela = 55.2200012207031
' .val_DemParcelas = 55.2200012207031
' .Desc_Plano = "JULHO/2007"
' .Val_RendaFatu = 0
' .LMI_Unico = 2
' .Per_Desc_LMI = 0
' .Num_Cotac_LMI = ""
' .Tip_Transmissao = 0
' .Ind_CondGerais = 1
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4383 |
' Set .ColecaoPedCorr = New Collection
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4384 |
' Set lobjstcPedCorr = New stcA46V705B
'    With lobjstcPedCorr
'         .NOSSO_NUMERO = ""
'         .Cod_Corr = 1601
'         .Num_Ped = 0
'         .Cod_Org = 0
'         .Cod_Ramo = 0
'         .Num_Apol = 0
'         .Num_Endo = 0
'         .Per_Corr = 10
'         .Num_Ordem_Corr = 1
'         .NUM_PI = 0
'    End With
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4385 |
' .ColecaoPedCorr.Add lobjstcPedCorr, lobjstcPedCorr.Class_ID
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4386 |
' Set .ColecaoPedCoss = New Collection
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4387 |
' Set .ColecaoPedJuros = New Collection
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4388 |
' Set lobjstcPedJuros = New stcA46V711B
'  With lobjstcPedJuros
' .NOSSO_NUMERO = "01133651010000016011"
' .Num_Parcela = 1
' .Val_Pos = 0
' .Val_Ant = 0
'  End With
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4389 |
'  .ColecaoPedJuros.Add lobjstcPedJuros, lobjstcPedJuros.Class_ID
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4390 |
' Set lobjstcPedJuros = New stcA46V711B
'  With lobjstcPedJuros
' .NOSSO_NUMERO = "01133651010000016011"
' .Num_Parcela = 2
' .Val_Pos = 0
' .Val_Ant = 0
'  End With
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4391 |
'  .ColecaoPedJuros.Add lobjstcPedJuros, lobjstcPedJuros.Class_ID
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4392 |
' Set lobjstcPedJuros = New stcA46V711B
'  With lobjstcPedJuros
' .NOSSO_NUMERO = "01133651010000016011"
' .Num_Parcela = 3
' .Val_Pos = 0
' .Val_Ant = 0
'  End With
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4393 |
'  .ColecaoPedJuros.Add lobjstcPedJuros, lobjstcPedJuros.Class_ID
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4394 |
' Set lobjstcPedJuros = New stcA46V711B
'  With lobjstcPedJuros
' .NOSSO_NUMERO = "01133651010000016011"
' .Num_Parcela = 4
' .Val_Pos = 0
' .Val_Ant = 0
'  End With
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4395 |
'  .ColecaoPedJuros.Add lobjstcPedJuros, lobjstcPedJuros.Class_ID
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4396 |
' Set lobjstcPedJuros = New stcA46V711B
'  With lobjstcPedJuros
' .NOSSO_NUMERO = "01133651010000016011"
' .Num_Parcela = 5
' .Val_Pos = 0
' .Val_Ant = 0
'  End With
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4397 |
'  .ColecaoPedJuros.Add lobjstcPedJuros, lobjstcPedJuros.Class_ID
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4398 |
' Set lobjstcPedJuros = New stcA46V711B
'  With lobjstcPedJuros
' .NOSSO_NUMERO = "01133651010000016011"
' .Num_Parcela = 6
' .Val_Pos = 0
' .Val_Ant = 0
'  End With
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4399 |
'  .ColecaoPedJuros.Add lobjstcPedJuros, lobjstcPedJuros.Class_ID
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4400 |
' Set lobjstcPedJuros = New stcA46V711B
'  With lobjstcPedJuros
' .NOSSO_NUMERO = "01133651010000016011"
' .Num_Parcela = 7
' .Val_Pos = 0
' .Val_Ant = 0
'  End With
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4401 |
'  .ColecaoPedJuros.Add lobjstcPedJuros, lobjstcPedJuros.Class_ID
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4402 |
' Set lobjstcPedJuros = New stcA46V711B
'  With lobjstcPedJuros
' .NOSSO_NUMERO = "01133651010000016011"
' .Num_Parcela = 8
' .Val_Pos = 0
' .Val_Ant = 0
'  End With
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4403 |
'  .ColecaoPedJuros.Add lobjstcPedJuros, lobjstcPedJuros.Class_ID
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4404 |
' Set lobjstcPedJuros = New stcA46V711B
'  With lobjstcPedJuros
' .NOSSO_NUMERO = "01133651010000016011"
' .Num_Parcela = 9
' .Val_Pos = 0
' .Val_Ant = 0
'  End With
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4405 |
'  .ColecaoPedJuros.Add lobjstcPedJuros, lobjstcPedJuros.Class_ID
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4406 |
' Set lobjstcPedJuros = New stcA46V711B
'  With lobjstcPedJuros
' .NOSSO_NUMERO = "01133651010000016011"
' .Num_Parcela = 10
' .Val_Pos = 0
' .Val_Ant = 0
'  End With
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4407 |
'  .ColecaoPedJuros.Add lobjstcPedJuros, lobjstcPedJuros.Class_ID
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4408 |
' Set lobjstcPedJuros = New stcA46V711B
'  With lobjstcPedJuros
' .NOSSO_NUMERO = "01133651010000016011"
' .Num_Parcela = 11
' .Val_Pos = 0
' .Val_Ant = 0
'  End With
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4409 |
'  .ColecaoPedJuros.Add lobjstcPedJuros, lobjstcPedJuros.Class_ID
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4410 |
' Set .ColecaoPedVenc = New Collection
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4411 |
' Set lobjstcPedParc = New stcA46V713B
'  With lobjstcPedParc
' .NOSSO_NUMERO = "01133651010000016011"
' .Cod_Org = 0
' .Cod_Ramo = 113
' .Num_Ped = 2
' .Num_Apol = 0
' .Num_Endo = 0
' .Num_Parc = 1
' .DAT_VENC = 20111226
'  End With
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4412 |
' .ColecaoPedVenc.Add lobjstcPedParc, lobjstcPedParc.Class_ID
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4413 |
' Set lobjstcPedParc = New stcA46V713B
'  With lobjstcPedParc
' .NOSSO_NUMERO = "01133651010000016011"
' .Cod_Org = 0
' .Cod_Ramo = 113
' .Num_Ped = 2
' .Num_Apol = 0
' .Num_Endo = 0
' .Num_Parc = 2
' .DAT_VENC = 20120126
'  End With
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4414 |
' .ColecaoPedVenc.Add lobjstcPedParc, lobjstcPedParc.Class_ID
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4415 |
' Set lobjstcPedParc = New stcA46V713B
'  With lobjstcPedParc
' .NOSSO_NUMERO = "01133651010000016011"
' .Cod_Org = 0
' .Cod_Ramo = 113
' .Num_Ped = 2
' .Num_Apol = 0
' .Num_Endo = 0
' .Num_Parc = 3
' .DAT_VENC = 20120226
'  End With
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4416 |
' .ColecaoPedVenc.Add lobjstcPedParc, lobjstcPedParc.Class_ID
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4417 |
' Set lobjstcPedParc = New stcA46V713B
'  With lobjstcPedParc
' .NOSSO_NUMERO = "01133651010000016011"
' .Cod_Org = 0
' .Cod_Ramo = 113
' .Num_Ped = 2
' .Num_Apol = 0
' .Num_Endo = 0
' .Num_Parc = 4
' .DAT_VENC = 20120326
'  End With
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4418 |
' .ColecaoPedVenc.Add lobjstcPedParc, lobjstcPedParc.Class_ID
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4419 |
' Set lobjstcPedParc = New stcA46V713B
'  With lobjstcPedParc
' .NOSSO_NUMERO = "01133651010000016011"
' .Cod_Org = 0
' .Cod_Ramo = 113
' .Num_Ped = 2
' .Num_Apol = 0
' .Num_Endo = 0
' .Num_Parc = 5
' .DAT_VENC = 20120426
'  End With
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4420 |
' .ColecaoPedVenc.Add lobjstcPedParc, lobjstcPedParc.Class_ID
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4421 |
' Set lobjstcPedParc = New stcA46V713B
'  With lobjstcPedParc
' .NOSSO_NUMERO = "01133651010000016011"
' .Cod_Org = 0
' .Cod_Ramo = 113
' .Num_Ped = 2
' .Num_Apol = 0
' .Num_Endo = 0
' .Num_Parc = 1
' .DAT_VENC = 20111226
'  End With
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4422 |
' .ColecaoPedVenc.Add lobjstcPedParc, lobjstcPedParc.Class_ID
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4423 |
' Set lobjstcPedParc = New stcA46V713B
'  With lobjstcPedParc
' .NOSSO_NUMERO = "01133651010000016011"
' .Cod_Org = 0
' .Cod_Ramo = 113
' .Num_Ped = 2
' .Num_Apol = 0
' .Num_Endo = 0
' .Num_Parc = 2
' .DAT_VENC = 20120126
'  End With
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4424 |
' .ColecaoPedVenc.Add lobjstcPedParc, lobjstcPedParc.Class_ID
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4425 |
' Set lobjstcPedParc = New stcA46V713B
'  With lobjstcPedParc
' .NOSSO_NUMERO = "01133651010000016011"
' .Cod_Org = 0
' .Cod_Ramo = 113
' .Num_Ped = 2
' .Num_Apol = 0
' .Num_Endo = 0
' .Num_Parc = 3
' .DAT_VENC = 20120226
'  End With
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4426 |
' .ColecaoPedVenc.Add lobjstcPedParc, lobjstcPedParc.Class_ID
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4427 |
' Set lobjstcPedParc = New stcA46V713B
'  With lobjstcPedParc
' .NOSSO_NUMERO = "01133651010000016011"
' .Cod_Org = 0
' .Cod_Ramo = 113
' .Num_Ped = 2
' .Num_Apol = 0
' .Num_Endo = 0
' .Num_Parc = 4
' .DAT_VENC = 20120326
'  End With
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4428 |
' .ColecaoPedVenc.Add lobjstcPedParc, lobjstcPedParc.Class_ID
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4429 |
' Set lobjstcPedParc = New stcA46V713B
'  With lobjstcPedParc
' .NOSSO_NUMERO = "01133651010000016011"
' .Cod_Org = 0
' .Cod_Ramo = 113
' .Num_Ped = 2
' .Num_Apol = 0
' .Num_Endo = 0
' .Num_Parc = 5
' .DAT_VENC = 20120426
'  End With
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4430 |
' .ColecaoPedVenc.Add lobjstcPedParc, lobjstcPedParc.Class_ID
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4431 |
' Set .objstcPedEndereco = New stcA46V707B
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4432 |
' With .objstcPedEndereco
' .NOSSO_NUMERO = "01133651010000016011"
' .Cod_Ramo = 113
' .Num_Apol = 0
' .Num_Endo = 0
' .Cod_Tip_Logr = "R"
' .Nom_Logr = "JOAO MORENO DOS SANTOS"
' .Dsc_Compl_Logr = "11"
' .Nom_Bairro = "VILA SOUZA"
' .Nom_Cid = "SAO PAULO"
' .Sig_UF = "SP"
' .Num_CEP = 2860120
' .Num_Fone = 8861133
' .NUM_PI = 1100000002
' .Cod_Reg_Cons = 0
' .Num_Reg_Cons = ""
' End With
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4433 |
' Set .objstcCtrlEmis = New stcA46V712B
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4434 |
' With .objstcCtrlEmis
' .NOSSO_NUMERO = "01133651010000016011"
' .NUM_PI = 1100000002
' .Num_Bloq = 0
' .COD_UNID = 251
' .Cod_Ramo = 113
' .Dat_Ini_Vig = 20111223
' .Dat_Ter_Vig = 20121223
' .Cod_Corr = 1601
' .Per_Corr = 10
' .Cod_Corr_2 = 0
' .Per_Corr_2 = 0
' .Cod_Corr_3 = 0
' .Per_Corr_3 = 0
' .Nom_Seg = "ANDRE LUIZ LOPES DE OLIVEIRA (YASUDA)"
' .Qtd_Anex = 1
' .Num_Apol_Ant = 11973275
' .Dsc_Obs_Prop = ""
' .Tip_Cotac = ""
' .Num_Cotac = ""
' .Per_Juros_Neg = 0
' .Cod_Envi_Doc = "X"
' .DAT_DIGIT = 20111202
' .HOR_DIGIT = 1416
' .Cod_Iden_Stat_Acao = 0
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4435 |
' .Cod_User_Digit = "YSNT21"
' .Dat_Digit_MF = 20111202
' .Hor_Digit_MF = 141600
' .Cod_User_Digit_MF = ""
' .Val_Pr_Tot = 276.119999542236
' .Cod_Iden_Classe = 0
' .Cod_Iden_Cobr = 0
' .Num_Prop = 2
' .Dat_Canc_Apol = 0
' .Cod_Tip_Doc = 3
' .Dat_Rec_Cia = 0
' .Num_Ref_Corr = ""
' .Dat_Emis = 0
' .Num_Apol = 0
' .Num_Endo = 0
' .Val_Pr_Emis = 207.14
' .Val_Pr_Liq = 207.14
' .Per_Tot_Coss = 0
' .Qtd_Parc = 5
' .Cod_Iden_Stat_Resp = 0
' .Cod_Iden_SAC = 0
' .Cod_Prod = 1692
' .Val_IS = 1692
' .Per_Desc_Coml = 0
' .Num_Cgc_Cpf = 28708974802#
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4436 |
' .Cod_Iden_Pessoa = 2
' .Cod_Iden_Pag_Antec = 0
' .DIG_BLOQUETO = 0
' End With
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4437 |
' Set .ColecaoTextoEspec = New Collection
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4438 |
' Set lobjstcPedEspec = New stcA46V726B
'  With lobjstcPedEspec
' .NOSSO_NUMERO = "01133651010000016011"
' .Cod_Espec = 8500
' .Cod_Org = 0
' .Cod_Ramo = 113
' .Num_Ped = 2
' .Num_Apol = 0
' .Num_Endo = 0
' .Num_pag = 1
' .Qtd_vias = 3
' .Espec_lin1 = ""
' .Espec_lin2 = ""
' .Espec_lin3 = ""
' .Espec_lin4 = ""
' .Espec_lin5 = ""
' .Espec_lin6 = ""
' .Espec_lin7 = ""
' .Espec_lin8 = ""
' .Espec_lin9 = ""
' .Espec_lin10 = ""
' .Espec_lin11 = ""
' .Espec_lin12 = ""
' .Espec_lin13 = ""
' .Espec_lin14 = ""
' .Espec_lin15 = ""
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4439 |
' .Espec_lin16 = ""
' .Espec_lin17 = ""
' .Espec_lin18 = ""
' .Espec_lin19 = ""
' .Espec_lin20 = "frmT46V155A"
' .NUM_PI = 1100000002
' End With
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4440 |
'.ColecaoTextoEspec.Add lobjstcPedEspec, lobjstcPedEspec.Class_ID
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4441 |
'  Set .ColecaoPedLoc = New Collection
' With lobjstcPedLoc
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4442 |
'  .NOSSO_NUMERO = ""
' .Num_Item = 1
' .Num_Ped = 0
' .Cod_Org = 0
' .Cod_Ramo = 0
' .Num_Apol = 0
' .Num_Endo = 0
' .Cod_Tip_Logr_Risco = "R"
' .Nom_Logr_Risco = "JOAO MORENO DOS SANTOS"
' .Dsc_Compl_Risco = ""
' .Nom_Bairro_Risco = "VILA SOUZA"
' .Nom_Cid_Risco = "SAO PAULO"
' .Sig_UF_Risco = "SP"
' .Num_Cep_Risco = 2860120
' .Per_Desc_Apar = 0
' .Per_Desc_Progres = 30
' .Per_IS_Vr = 100
' .Val_Coef_IS_Vr = 0
' .Val_Vr = 250000
' .Tab_Clasf = 0
' .Cod_Clasf = "001"
' .Cod_Compl_Clasf = "11"
' .Dsc_Clasf = ""
' .Cod_Classe_Ocup = 0
' .Cod_Constr = 1
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4443 |
' .Cod_User_Digit = ""
' .DAT_DIGIT = 0
' .HOR_DIGIT = 0
' .Cod_User_Liber = ""
' .Dat_Liber = 0
' .Hor_Liber = 0
' .Cod_Sit_item = 0
' .Per_Bonus = 10
' .Qtd_Empreg = 0
' .Num_Insp = 0
' .Cod_Iden_Segur_01_V = ""
' .Cod_Iden_Segur_02_A = ""
' .Cod_Iden_Segur_03_P = ""
' .Cod_Iden_Segur_04_I = ""
' .Cod_Iden_Segur_05_G = ""
' .Cod_Iden_Segur_06 = ""
' .Cod_Iden_Segur_07 = ""
' .Cod_Iden_Segur_08 = ""
' .Cod_Iden_Segur_09 = ""
' .Cod_Iden_Segur_10 = ""
' .Per_Tot_Segur = 0
' .Cod_Iden_Protec_01_E = ""
' .Cod_Iden_Protec_02_H = ""
' .Cod_Iden_Protec_03_S = ""
' .Cod_Iden_Protec_04_D = ""
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4444 |
' .Cod_Iden_Protec_05 = ""
' .Cod_Iden_Protec_06 = ""
' .Cod_Iden_Protec_07 = ""
' .Cod_Iden_Protec_08 = ""
' .Cod_Iden_Protec_09 = ""
' .Cod_Iden_Protec_10 = ""
' .Per_Tot_Protec = 0
' .Cod_Produto = 0
' .Cod_Regiao = 0
' .Cod_Iden_Tarif = 0
' .NUM_PI = 0
' .Num_Item_Ant = 1
' .Tip_Emissao = 200
' .Num_Apol_Anterior = 11973275
' .Num_Item_Anterior = 0
' .Ind_Sinistro = 0
' .Insp_Obr = False
' .Insp_Agend = False
' .Dig_Bens = False
' .Cob_Bens = False
' .Tip_Atividade = 0
' .Ren_Con_Seguradora = 0
' .Ren_Con_Apolice = ""
' .Ren_Con_Item = ""
' .Tip_Cobert = 0
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4445 |
' .Cod_Plano = 1
' Set .ColecaoPedCobert = New Collection
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4446 |
' Set lobjstcPedCobert = New stcA46V704B
' With lobjstcPedCobert
' .NOSSO_NUMERO = ""
' .Cod_Cobert = 1
' .Num_Ped = 0
' .Cod_Org = 0
' .Cod_Ramo = 0
' .Num_Apol = 0
' .Num_Endo = 0
' .Num_Item = 0
' .Tab_Cobert = 0
' .Val_IS = 250000
' .Per_Taxa = 3.08999996632338E-02
' .Per_Desc_Coml = 30
' .Per_Desc_Exp = 0
' .VAL_PR = 0
' .Val_Pr_Anu = 0
' .Val_Pr_Tarif = 0
' .Val_Pr_Tarif_Anu = 0
' .Val_Pr_Resg = 0
' .Val_Pr_Resg_Anu = 0
' .Per_Franq = 0
' .Val_Franq_Min = 0
' .Val_Franq_Max = 0
' .NUM_PI = 0
' .Num_Meses_PI = 0
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4447 |
' .Formula = ""
' .COD_CLASSE = 0
' .Val_Fator_Franq = 1
' .Alterar_Cobert = ""
' .Val_Risco_Cob = 250000
' .Nom_Cobertura = "INCENDIO / QUEDA DE RAIO / EXPLOSAO"
' .Cod_Grupo_Cobertura = 0
' End With
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4448 |
' .ColecaoPedCobert.Add lobjstcPedCobert, lobjstcPedCobert.Class_ID
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4449 |
' Set lobjstcPedCobert = New stcA46V704B
' With lobjstcPedCobert
' .NOSSO_NUMERO = ""
' .Cod_Cobert = 4
' .Num_Ped = 0
' .Cod_Org = 0
' .Cod_Ramo = 0
' .Num_Apol = 0
' .Num_Endo = 0
' .Num_Item = 0
' .Tab_Cobert = 0
' .Val_IS = 1000
' .Per_Taxa = 3
' .Per_Desc_Coml = 30
' .Per_Desc_Exp = 0
' .VAL_PR = 0
' .Val_Pr_Anu = 0
' .Val_Pr_Tarif = 0
' .Val_Pr_Tarif_Anu = 0
' .Val_Pr_Resg = 0
' .Val_Pr_Resg_Anu = 0
' .Per_Franq = 0
' .Val_Franq_Min = 0
' .Val_Franq_Max = 0
' .NUM_PI = 0
' .Num_Meses_PI = 0
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4450 |
' .Formula = ""
' .COD_CLASSE = 0
' .Val_Fator_Franq = 1
' .Alterar_Cobert = ""
' .Val_Risco_Cob = 1000
' .Nom_Cobertura = "VIDROS E ESPELHOS"
' .Cod_Grupo_Cobertura = 0
' End With
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4451 |
' .ColecaoPedCobert.Add lobjstcPedCobert, lobjstcPedCobert.Class_ID
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4452 |
' Set lobjstcPedCobert = New stcA46V704B
' With lobjstcPedCobert
' .NOSSO_NUMERO = ""
' .Cod_Cobert = 5
' .Num_Ped = 0
' .Cod_Org = 0
' .Cod_Ramo = 0
' .Num_Apol = 0
' .Num_Endo = 0
' .Num_Item = 0
' .Tab_Cobert = 0
' .Val_IS = 6000
' .Per_Taxa = 2.57599997520447
' .Per_Desc_Coml = 30
' .Per_Desc_Exp = 0
' .VAL_PR = 0
' .Val_Pr_Anu = 0
' .Val_Pr_Tarif = 0
' .Val_Pr_Tarif_Anu = 0
' .Val_Pr_Resg = 0
' .Val_Pr_Resg_Anu = 0
' .Per_Franq = 0
' .Val_Franq_Min = 0
' .Val_Franq_Max = 0
' .NUM_PI = 0
' .Num_Meses_PI = 0
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4453 |
' .Formula = ""
' .COD_CLASSE = 0
' .Val_Fator_Franq = 1
' .Alterar_Cobert = ""
' .Val_Risco_Cob = 6000
' .Nom_Cobertura = "ROUBO OU FURTO COM VESTIGIOS"
' .Cod_Grupo_Cobertura = 0
' End With
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4454 |
' .ColecaoPedCobert.Add lobjstcPedCobert, lobjstcPedCobert.Class_ID
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4455 |
' Set lobjstcPedCobert = New stcA46V704B
' With lobjstcPedCobert
' .NOSSO_NUMERO = ""
' .Cod_Cobert = 6
' .Num_Ped = 0
' .Cod_Org = 0
' .Cod_Ramo = 0
' .Num_Apol = 0
' .Num_Endo = 0
' .Num_Item = 0
' .Tab_Cobert = 0
' .Val_IS = 30000
' .Per_Taxa = 7.99999982118607E-02
' .Per_Desc_Coml = 30
' .Per_Desc_Exp = 0
' .VAL_PR = 0
' .Val_Pr_Anu = 0
' .Val_Pr_Tarif = 0
' .Val_Pr_Tarif_Anu = 0
' .Val_Pr_Resg = 0
' .Val_Pr_Resg_Anu = 0
' .Per_Franq = 0
' .Val_Franq_Min = 0
' .Val_Franq_Max = 0
' .NUM_PI = 0
' .Num_Meses_PI = 0
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4456 |
' .Formula = ""
' .COD_CLASSE = 0
' .Val_Fator_Franq = 1
' .Alterar_Cobert = ""
' .Val_Risco_Cob = 30000
' .Nom_Cobertura = "RESPONSABILIDADE CIVIL FAMILIAR"
' .Cod_Grupo_Cobertura = 0
' End With
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4457 |
' .ColecaoPedCobert.Add lobjstcPedCobert, lobjstcPedCobert.Class_ID
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4458 |
' Set lobjstcPedCobert = New stcA46V704B
' With lobjstcPedCobert
' .NOSSO_NUMERO = ""
' .Cod_Cobert = 12
' .Num_Ped = 0
' .Cod_Org = 0
' .Cod_Ramo = 0
' .Num_Apol = 0
' .Num_Endo = 0
' .Num_Item = 0
' .Tab_Cobert = 0
' .Val_IS = 5000
' .Per_Taxa = 0.925499975681305
' .Per_Desc_Coml = 30
' .Per_Desc_Exp = 0
' .VAL_PR = 0
' .Val_Pr_Anu = 0
' .Val_Pr_Tarif = 0
' .Val_Pr_Tarif_Anu = 0
' .Val_Pr_Resg = 0
' .Val_Pr_Resg_Anu = 0
' .Per_Franq = 10
' .Val_Franq_Min = 200
' .Val_Franq_Max = 0
' .NUM_PI = 0
' .Num_Meses_PI = 0
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4459 |
' .Formula = ""
' .COD_CLASSE = 0
' .Val_Fator_Franq = 1
' .Alterar_Cobert = ""
' .Val_Risco_Cob = 5000
' .Nom_Cobertura = "DANOS ELETRICOS"
' .Cod_Grupo_Cobertura = 0
' End With
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4460 |
' .ColecaoPedCobert.Add lobjstcPedCobert, lobjstcPedCobert.Class_ID
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4461 |
' Set lobjstcPedCobert = New stcA46V704B
' With lobjstcPedCobert
' .NOSSO_NUMERO = ""
' .Cod_Cobert = 18
' .Num_Ped = 0
' .Cod_Org = 0
' .Cod_Ramo = 0
' .Num_Apol = 0
' .Num_Endo = 0
' .Num_Item = 0
' .Tab_Cobert = 0
' .Val_IS = 5000
' .Per_Taxa = 0.230000004172325
' .Per_Desc_Coml = 30
' .Per_Desc_Exp = 0
' .VAL_PR = 0
' .Val_Pr_Anu = 0
' .Val_Pr_Tarif = 0
' .Val_Pr_Tarif_Anu = 0
' .Val_Pr_Resg = 0
' .Val_Pr_Resg_Anu = 0
' .Per_Franq = 10
' .Val_Franq_Min = 200
' .Val_Franq_Max = 0
' .NUM_PI = 0
' .Num_Meses_PI = 0
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4462 |
' .Formula = ""
' .COD_CLASSE = 0
' .Val_Fator_Franq = 1
' .Alterar_Cobert = ""
' .Val_Risco_Cob = 5000
' .Nom_Cobertura = "VENDAVAL"
' .Cod_Grupo_Cobertura = 0
' End With
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4463 |
' .ColecaoPedCobert.Add lobjstcPedCobert, lobjstcPedCobert.Class_ID
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4464 |
' Set lobjstcPedCobert = New stcA46V704B
' With lobjstcPedCobert
' .NOSSO_NUMERO = ""
' .Cod_Cobert = 27
' .Num_Ped = 0
' .Cod_Org = 0
' .Cod_Ramo = 0
' .Num_Apol = 0
' .Num_Endo = 0
' .Num_Item = 0
' .Tab_Cobert = 0
' .Val_IS = 1500
' .Per_Taxa = 5
' .Per_Desc_Coml = 30
' .Per_Desc_Exp = 0
' .VAL_PR = 0
' .Val_Pr_Anu = 0
' .Val_Pr_Tarif = 0
' .Val_Pr_Tarif_Anu = 0
' .Val_Pr_Resg = 0
' .Val_Pr_Resg_Anu = 0
' .Per_Franq = 0
' .Val_Franq_Min = 250
' .Val_Franq_Max = 0
' .NUM_PI = 0
' .Num_Meses_PI = 0
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4465 |
' .Formula = ""
' .COD_CLASSE = 0
' .Val_Fator_Franq = 1
' .Alterar_Cobert = ""
' .Val_Risco_Cob = 1500
' .Nom_Cobertura = "TODOS OS RISCOS (BRASIL) - SEM EXTORSAO"
' .Cod_Grupo_Cobertura = 0
' End With
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4466 |
' .ColecaoPedCobert.Add lobjstcPedCobert, lobjstcPedCobert.Class_ID
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4467 |
' Set .ColecaoPedLocBem = New Collection
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4468 |
' Set lobjstcPedBem = New stcA46V715B
' With lobjstcPedBem
' .NOSSO_NUMERO = ""
' .Num_Ped = 0
' .Cod_Org = 0
' .Cod_Ramo = 0
' .Num_Apol = 0
' .Num_Endo = 0
' .Num_Item = 0
' .Dsc_Bem = "CALOI"
' .Val_Bem = 1000
' .Cod_Bem = 7
' .Cod_Seq = 1
' End With
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4469 |
' .ColecaoPedLocBem.Add lobjstcPedBem, lobjstcPedBem.Class_ID
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4470 |
' Set .ColecaoPedSeguro = New Collection
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4471 |
' Set .ColecaoPedFator = New Collection
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4472 |
' Set .objstcPedObs = New stcA46V714B
' With .objstcPedObs
' .NOSSO_NUMERO = ""
' .Num_Item = 0
' .Observacao = ""
' End With
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4473 |
' Set .objstcPedInsp = New stcA46V717B
' With .objstcPedInsp
' .NOSSO_NUMERO = ""
' .Num_Ped = 0
' .Cod_Org = 0
' .Cod_Ramo = 0
' .Num_Item = 0
' .Num_Insp = 0
' .Dat_Cadastro = 0
' .Pess_Cont = ""
' .Num_Tel = ""
' .Num_Logr_Risco = 11
' End With
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4474 |
' Set .ColecaoPedClau = New Collection
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4475 |
' Set .ColecaoNormas = New Collection
'
''02/12/2011 | | 14:16:55 | XXXXXXXXXXXXXXXXXXXX | WEBSYAS          | D46O003      | 4476 |
' End With
' .ColecaoPedLoc.Add lobjstcPedLoc, lobjstcPedLoc.Class_ID
'
'End With
'
'Call a.gfSalvar(lobjstcPed)
'
'    Call a.gfCalcular("01133651010000016011", lColecaoA, lColecaoB, lColecaoC, lClasse4)
'    'Set objstc = a.gfProRenYasuda("0011747805", 1, True)
'    'Call a.gfCancelarEmissao("01127854204000493010", lRet)
' '   a.gfSeguroNovo (112)
'  '  Call a.gfSolicitarClasse4("01127854204000493010", "tjgomes@yasuda.com.br", "(11) 9999-9999", "solicitou a atenção para esta cotação", "thais", "erro de desconto acima do limite")
'   End
'End Sub
'
'Public Function teste()
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0031 |
' Dim lobjstcPedLoc   As stcA46V708B
' Dim lObjstcPedClau  As stcA46V703B
' Dim lobjstcPedCobert  As stcA46V704B
' Dim lobjstcPedCorr  As stcA46V705B
' Dim lobjstcPedCoss  As stcA46V706B
' Dim lobjstcPedEnder  As stcA46V707B
' Dim lobjstcPedJuros  As stcA46V711B
' Dim lobjstcPedCtrEmis As stcA46V712B
' Dim lobjstcPedParc As stcA46V713B
' Dim lobjstcPedObs As stcA46V714B
' Dim lobjstcPedBem As stcA46V715B
' Dim lobjstcPedOutroSeg As stcA46V716B
' Dim lobjstcPedInsp As stcA46V717B
' Dim lobjstcFator    As stcA46V720B
' Dim lobjstcPedEspec As stcA46V726B
' Dim lstcNorma       As stcA46V708N
' Dim lobjstcPed      As stcA46V702B
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0032 |
'
' Set lobjstcPed = New stcA46V702B
'    With lobjstcPed
' .NOSSO_NUMERO = "01129031531009682261"
' .Cod_Ramo = 114
' .Cod_Org = 0
' .Num_Ped = 212
' .Num_Apol = 0
' .Num_Endo = 0
' .Num_Apol_An = 11411271
' .Cod_Reg = 1
' .Cod_Bco_Cobr = 341
' .Cod_Agen_Angar = 0
' .Cod_Agen_Indicr = 0
' .Num_Cta_Indicr = 0
' .Cod_Unid_Prod = 6631
' .Cod_Prod = 4265
' .Cod_Agen_Bas_Cor = 0
' .Cod_Dig_Agen_Bas_Cor = 0
' .Num_Cta_Bas_Cor = 0
' .Cod_Segurado = 0
' .Dig_Cod_Segurado = ""
' .Nom_Segurado = "COND. EDIF. SYBELLE"
' .Dsc_Compl_Segurado = ""
' .Num_Cgc_Cpf = 1846042000100#
' .Cod_Tip_Logr = "R"
' .Nom_Logr = "ANA BERTA ROSKAMP"
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0033 |
' .Dsc_Compl_Logr = "55"
' .Nom_Bairro = "JARDIM DAS AMERICAS"
' .Nom_Cid = "CURITIBA"
' .Sig_UF = "PR"
' .Num_CEP = 81530250
' .Num_Fone = 33666410
' .Cod_Estip = 1
' .Val_Coef_Cms_Pad = 0
' .Dat_Ini_Vig = 20111031
' .Dat_Ter_Vig = 20121031
' .Dat_Emis = 0
' .Cod_Moed = 0
' .Per_Juro_Frac = 0
' .Qtd_Parc = 4
' .Qtd_Anexo = 1
' .Cod_Iden_Praz = 1
' .Cod_Pr_Min_Endo = 1
' .Cod_Pr_Min_Parc = 1
' .Val_Pr_Antec = 0
' .Per_Adic_Form_Pag = 0
' .Per_Desc_Form_Pag = 0
' .Val_IS = 2400000
' .Val_IS_2 = 0
' .VAL_PR = 274.86
' .Val_Pr_Ref_Tarif = 274.86
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0034 |
' .Val_Pr_Ref_Resg = 274.86
' .Dat_Ini_Viag = 0
' .Cod_Tip_Liber = 0
' .Cod_Iden_Alt_Corr = ""
' .Cod_User_Digit = ""
' .DAT_DIGIT = 20111031
' .HOR_DIGIT = 844
' .Cod_User_Liber = "968226"
' .Dat_Liber = 20111031
' .Hor_Liber = 844
' .Cod_Sit_Ped = 3
' .Cod_Tip_Doc = 2
' .Cod_Dizer = 0
' .Qtd_Copia = 3
' .Val_Cust_Emis = 0
' .Val_Parc_Min = 0
' .Val_Juro = 0
' .Val_IOF = 27.66
' .Num_Prop = 0
' .Cod_Posi_Atual = ""
' .Dat_Mov_Ini = 0
' .Dat_Mov_Ter = 0
' .Tip_Cust_Emis = 1
' .Val_Cambio_Emis = 0
' .Iden = ""
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0035 |
' .NUM_PI = 1100000212
' .Cod_Produto = 1018
' .Cod_Protocolo = 0
' .Cod_Sequencia = 0
' .Tip_Pessoa = 1
' .Cod_Ramo_Ativ = 15
' .Cod_Doc = 0
' .Num_Doc = ""
' .Org_Exp = ""
' .Dat_Emis_Doc = 0
' .Nom_Empresa = ""
' .Sexo = ""
' .Dat_Pri_Calculo = 20111017
' .Ult_Tarifa = 10180
' .Cod_Unicidade = 69369
' .Cod_Ver_Unici = 1
' .Val_PriParcela = 68.7099990844727
' .val_DemParcelas = 68.7099990844727
' .Desc_Plano = ""
' .Val_RendaFatu = 0
' .LMI_Unico = 0
' .Per_Desc_LMI = 0
' .Num_Cotac_LMI = ""
' .Tip_Transmissao = 0
' .Ind_CondGerais = 1
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0036 |
' Set .ColecaoPedCorr = New Collection
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0037 |
' Set lobjstcPedCorr = New stcA46V705B
'    With lobjstcPedCorr
'         .NOSSO_NUMERO = "01129031531009682261"
'         .Cod_Corr = 1601
'         .Num_Ped = 212
'         .Cod_Org = 0
'         .Cod_Ramo = 114
'         .Num_Apol = 0
'         .Num_Endo = 0
'         .Per_Corr = 15
'         .Num_Ordem_Corr = 1
'         .NUM_PI = 1100000212
'    End With
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0038 |
' .ColecaoPedCorr.Add lobjstcPedCorr, lobjstcPedCorr.Class_ID
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0039 |
' Set .ColecaoPedCoss = New Collection
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0040 |
' Set .ColecaoPedJuros = New Collection
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0041 |
' Set lobjstcPedJuros = New stcA46V711B
'  With lobjstcPedJuros
' .NOSSO_NUMERO = "01129031531009682261"
' .Num_Parcela = 1
' .Val_Pos = 0
' .Val_Ant = 0
'  End With
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0042 |
'  .ColecaoPedJuros.Add lobjstcPedJuros, lobjstcPedJuros.Class_ID
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0043 |
' Set lobjstcPedJuros = New stcA46V711B
'  With lobjstcPedJuros
' .NOSSO_NUMERO = "01129031531009682261"
' .Num_Parcela = 2
' .Val_Pos = 0
' .Val_Ant = 0
'  End With
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0044 |
'  .ColecaoPedJuros.Add lobjstcPedJuros, lobjstcPedJuros.Class_ID
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0045 |
' Set lobjstcPedJuros = New stcA46V711B
'  With lobjstcPedJuros
' .NOSSO_NUMERO = "01129031531009682261"
' .Num_Parcela = 3
' .Val_Pos = 0
' .Val_Ant = 0
'  End With
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0046 |
'  .ColecaoPedJuros.Add lobjstcPedJuros, lobjstcPedJuros.Class_ID
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0047 |
' Set lobjstcPedJuros = New stcA46V711B
'  With lobjstcPedJuros
' .NOSSO_NUMERO = "01129031531009682261"
' .Num_Parcela = 4
' .Val_Pos = 0
' .Val_Ant = 0
'  End With
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0048 |
'  .ColecaoPedJuros.Add lobjstcPedJuros, lobjstcPedJuros.Class_ID
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0049 |
' Set lobjstcPedJuros = New stcA46V711B
'  With lobjstcPedJuros
' .NOSSO_NUMERO = "01129031531009682261"
' .Num_Parcela = 5
' .Val_Pos = 1.5
' .Val_Ant = 0
'  End With
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0050 |
'  .ColecaoPedJuros.Add lobjstcPedJuros, lobjstcPedJuros.Class_ID
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0051 |
' Set lobjstcPedJuros = New stcA46V711B
'  With lobjstcPedJuros
' .NOSSO_NUMERO = "01129031531009682261"
' .Num_Parcela = 6
' .Val_Pos = 1.5
' .Val_Ant = 1.5
'  End With
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0052 |
'  .ColecaoPedJuros.Add lobjstcPedJuros, lobjstcPedJuros.Class_ID
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0053 |
' Set lobjstcPedJuros = New stcA46V711B
'  With lobjstcPedJuros
' .NOSSO_NUMERO = "01129031531009682261"
' .Num_Parcela = 7
' .Val_Pos = 1.5
' .Val_Ant = 1.5
'  End With
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0054 |
'  .ColecaoPedJuros.Add lobjstcPedJuros, lobjstcPedJuros.Class_ID
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0055 |
' Set lobjstcPedJuros = New stcA46V711B
'  With lobjstcPedJuros
' .NOSSO_NUMERO = "01129031531009682261"
' .Num_Parcela = 8
' .Val_Pos = 1.5
' .Val_Ant = 1.5
'  End With
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0056 |
'  .ColecaoPedJuros.Add lobjstcPedJuros, lobjstcPedJuros.Class_ID
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0057 |
' Set lobjstcPedJuros = New stcA46V711B
'  With lobjstcPedJuros
' .NOSSO_NUMERO = "01129031531009682261"
' .Num_Parcela = 9
' .Val_Pos = 1.5
' .Val_Ant = 1.5
'  End With
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0058 |
'  .ColecaoPedJuros.Add lobjstcPedJuros, lobjstcPedJuros.Class_ID
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0059 |
' Set lobjstcPedJuros = New stcA46V711B
'  With lobjstcPedJuros
' .NOSSO_NUMERO = "01129031531009682261"
' .Num_Parcela = 10
' .Val_Pos = 1.5
' .Val_Ant = 1.5
'  End With
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0060 |
'  .ColecaoPedJuros.Add lobjstcPedJuros, lobjstcPedJuros.Class_ID
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0061 |
' Set lobjstcPedJuros = New stcA46V711B
'  With lobjstcPedJuros
' .NOSSO_NUMERO = "01129031531009682261"
' .Num_Parcela = 11
' .Val_Pos = 1.5
' .Val_Ant = 1.5
'  End With
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0062 |
'  .ColecaoPedJuros.Add lobjstcPedJuros, lobjstcPedJuros.Class_ID
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0063 |
' Set .ColecaoPedVenc = New Collection
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0064 |
' Set lobjstcPedParc = New stcA46V713B
'  With lobjstcPedParc
' .NOSSO_NUMERO = "01129031531009682261"
' .Cod_Org = 0
' .Cod_Ramo = 114
' .Num_Ped = 212
' .Num_Apol = 0
' .Num_Endo = 0
' .Num_Parc = 1
' .DAT_VENC = 20111201
'  End With
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0065 |
' .ColecaoPedVenc.Add lobjstcPedParc, lobjstcPedParc.Class_ID
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0066 |
' Set lobjstcPedParc = New stcA46V713B
'  With lobjstcPedParc
' .NOSSO_NUMERO = "01129031531009682261"
' .Cod_Org = 0
' .Cod_Ramo = 114
' .Num_Ped = 212
' .Num_Apol = 0
' .Num_Endo = 0
' .Num_Parc = 2
' .DAT_VENC = 20120101
'  End With
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0067 |
' .ColecaoPedVenc.Add lobjstcPedParc, lobjstcPedParc.Class_ID
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0068 |
' Set lobjstcPedParc = New stcA46V713B
'  With lobjstcPedParc
' .NOSSO_NUMERO = "01129031531009682261"
' .Cod_Org = 0
' .Cod_Ramo = 114
' .Num_Ped = 212
' .Num_Apol = 0
' .Num_Endo = 0
' .Num_Parc = 3
' .DAT_VENC = 20120201
'  End With
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0069 |
' .ColecaoPedVenc.Add lobjstcPedParc, lobjstcPedParc.Class_ID
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0070 |
' Set lobjstcPedParc = New stcA46V713B
'  With lobjstcPedParc
' .NOSSO_NUMERO = "01129031531009682261"
' .Cod_Org = 0
' .Cod_Ramo = 114
' .Num_Ped = 212
' .Num_Apol = 0
' .Num_Endo = 0
' .Num_Parc = 4
' .DAT_VENC = 20120301
'  End With
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0071 |
' .ColecaoPedVenc.Add lobjstcPedParc, lobjstcPedParc.Class_ID
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0072 |
' Set .objstcPedEndereco = New stcA46V707B
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0073 |
' With .objstcPedEndereco
' .NOSSO_NUMERO = "01129031531009682261"
' .Cod_Ramo = 114
' .Num_Apol = 0
' .Num_Endo = 0
' .Cod_Tip_Logr = "R"
' .Nom_Logr = "XV DE NOVEMBRO"
' .Dsc_Compl_Logr = "2535"
' .Nom_Bairro = "ALTO DA RUA XV"
' .Nom_Cid = "CURITIBA"
' .Sig_UF = "PR"
' .Num_CEP = 80050000
' .Num_Fone = 99559306
' .NUM_PI = 1100000212
' .Cod_Reg_Cons = 0
' .Num_Reg_Cons = ""
' End With
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0074 |
' Set .objstcCtrlEmis = New stcA46V712B
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0075 |
' With .objstcCtrlEmis
' .NOSSO_NUMERO = "01129031531009682261"
' .NUM_PI = 1100000212
' .Num_Bloq = 0
' .COD_UNID = 6631
' .Cod_Ramo = 114
' .Dat_Ini_Vig = 20111031
' .Dat_Ter_Vig = 20121031
' .Cod_Corr = 1601
' .Per_Corr = 15
' .Cod_Corr_2 = 0
' .Per_Corr_2 = 0
' .Cod_Corr_3 = 0
' .Per_Corr_3 = 0
' .Nom_Seg = "COND. EDIF. SYBELLE"
' .Qtd_Anex = 0
' .Num_Apol_Ant = 11411271
' .Dsc_Obs_Prop = "0"
' .Tip_Cotac = "0"
' .Num_Cotac = "000000000000"
' .Per_Juros_Neg = 0
' .Cod_Envi_Doc = "X"
' .DAT_DIGIT = 20111017
' .HOR_DIGIT = 84800
' .Cod_Iden_Stat_Acao = 0
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0076 |
' .Cod_User_Digit = ""
' .Dat_Digit_MF = 20111017
' .Hor_Digit_MF = 84400
' .Cod_User_Digit_MF = ""
' .Val_Pr_Tot = 402
' .Cod_Iden_Classe = 0
' .Cod_Iden_Cobr = 0
' .Num_Prop = 212
' .Dat_Canc_Apol = 0
' .Cod_Tip_Doc = 3
' .Dat_Rec_Cia = 0
' .Num_Ref_Corr = "0"
' .Dat_Emis = 0
' .Num_Apol = 0
' .Num_Endo = 0
' .Val_Pr_Emis = 274
' .Val_Pr_Liq = 274
' .Per_Tot_Coss = 0
' .Qtd_Parc = 0
' .Cod_Iden_Stat_Resp = 0
' .Cod_Iden_SAC = 0
' .Cod_Prod = 4265
' .Val_IS = 4265
' .Per_Desc_Coml = 0
' .Num_Cgc_Cpf = 1846042000100#
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0077 |
' .Cod_Iden_Pessoa = 1
' .Cod_Iden_Pag_Antec = 2
' .DIG_BLOQUETO = 0
' End With
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0078 |
' Set .ColecaoTextoEspec = New Collection
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0079 |
'  Set .ColecaoPedLoc = New Collection
' With lobjstcPedLoc
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0080 |
'  .NOSSO_NUMERO = "01129031531009682261"
' .Num_Item = 1
' .Num_Ped = 212
' .Cod_Org = 0
' .Cod_Ramo = 114
' .Num_Apol = 0
' .Num_Endo = 0
' .Cod_Tip_Logr_Risco = "R"
' .Nom_Logr_Risco = "XV DE NOVEMBRO"
' .Dsc_Compl_Risco = ""
' .Nom_Bairro_Risco = "ALTO DA RUA XV"
' .Nom_Cid_Risco = "CURITIBA"
' .Sig_UF_Risco = "PR"
' .Num_Cep_Risco = 80050000
' .Per_Desc_Apar = 0
' .Per_Desc_Progres = 7
' .Per_IS_Vr = 100
' .Val_Coef_IS_Vr = 0
' .Val_Vr = 2400000
' .Tab_Clasf = 1
' .Cod_Clasf = "001"
' .Cod_Compl_Clasf = "11"
' .Dsc_Clasf = "CONDOMINIO RESIDENCIAL"
' .Cod_Classe_Ocup = 0
' .Cod_Constr = 0
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0081 |
' .Cod_User_Digit = "968226"
' .DAT_DIGIT = 20111031
' .HOR_DIGIT = 844
' .Cod_User_Liber = "968226"
' .Dat_Liber = 20111031
' .Hor_Liber = 844
' .Cod_Sit_item = 1
' .Per_Bonus = 5
' .Qtd_Empreg = 0
' .Num_Insp = 0
' .Cod_Iden_Segur_01_V = "N"
' .Cod_Iden_Segur_02_A = "N"
' .Cod_Iden_Segur_03_P = "N"
' .Cod_Iden_Segur_04_I = "N"
' .Cod_Iden_Segur_05_G = "N"
' .Cod_Iden_Segur_06 = "N"
' .Cod_Iden_Segur_07 = "N"
' .Cod_Iden_Segur_08 = "N"
' .Cod_Iden_Segur_09 = "N"
' .Cod_Iden_Segur_10 = "N"
' .Per_Tot_Segur = 0
' .Cod_Iden_Protec_01_E = "N"
' .Cod_Iden_Protec_02_H = "N"
' .Cod_Iden_Protec_03_S = "N"
' .Cod_Iden_Protec_04_D = "N"
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0082 |
' .Cod_Iden_Protec_05 = "N"
' .Cod_Iden_Protec_06 = "N"
' .Cod_Iden_Protec_07 = "N"
' .Cod_Iden_Protec_08 = "N"
' .Cod_Iden_Protec_09 = "N"
' .Cod_Iden_Protec_10 = "N"
' .Per_Tot_Protec = 0
' .Cod_Produto = 1
' .Cod_Regiao = 70
' .Cod_Iden_Tarif = 1018
' .NUM_PI = 1100000212
' .Num_Item_Ant = 1
' .Tip_Emissao = 200
' .Num_Apol_Anterior = 11411271
' .Num_Item_Anterior = 1
' .Ind_Sinistro = 2
' .Insp_Obr = False
' .Insp_Agend = False
' .Dig_Bens = False
' .Cob_Bens = False
' .Tip_Atividade = 0
' .Ren_Con_Seguradora = 0
' .Ren_Con_Apolice = ""
' .Ren_Con_Item = ""
' .Tip_Cobert = 0
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0083 |
' .Cod_Plano = 1
' Set .ColecaoPedCobert = New Collection
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0084 |
' Set lobjstcPedCobert = New stcA46V704B
' With lobjstcPedCobert
' .NOSSO_NUMERO = "01129031531009682261"
' .Cod_Cobert = 1
' .Num_Ped = 212
' .Cod_Org = 0
' .Cod_Ramo = 114
' .Num_Apol = 0
' .Num_Endo = 0
' .Num_Item = 1
' .Tab_Cobert = 1
' .Val_IS = 2400000
' .Per_Taxa = 0.0089
' .Per_Desc_Coml = 0
' .Per_Desc_Exp = 25
' .VAL_PR = 141.12
' .Val_Pr_Anu = 141.12
' .Val_Pr_Tarif = 141.12
' .Val_Pr_Tarif_Anu = 141.12
' .Val_Pr_Resg = 141.12
' .Val_Pr_Resg_Anu = 141.12
' .Per_Franq = 10
' .Val_Franq_Min = 1000
' .Val_Franq_Max = 0
' .NUM_PI = 1100000212
' .Num_Meses_PI = 0
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0085 |
' .Formula = ""
' .COD_CLASSE = 1
' .Val_Fator_Franq = 0
' .Alterar_Cobert = ""
' .Val_Risco_Cob = 2400000
' .Nom_Cobertura = ""
' .Cod_Grupo_Cobertura = 0
' End With
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0086 |
' .ColecaoPedCobert.Add lobjstcPedCobert, lobjstcPedCobert.Class_ID
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0087 |
' Set lobjstcPedCobert = New stcA46V704B
' With lobjstcPedCobert
' .NOSSO_NUMERO = "01129031531009682261"
' .Cod_Cobert = 4
' .Num_Ped = 212
' .Cod_Org = 0
' .Cod_Ramo = 114
' .Num_Apol = 0
' .Num_Endo = 0
' .Num_Item = 1
' .Tab_Cobert = 1
' .Val_IS = 2000
' .Per_Taxa = 2.4
' .Per_Desc_Coml = 0
' .Per_Desc_Exp = 25
' .VAL_PR = 29.54
' .Val_Pr_Anu = 29.54
' .Val_Pr_Tarif = 29.54
' .Val_Pr_Tarif_Anu = 29.54
' .Val_Pr_Resg = 29.54
' .Val_Pr_Resg_Anu = 29.54
' .Per_Franq = 0
' .Val_Franq_Min = 100
' .Val_Franq_Max = 0
' .NUM_PI = 1100000212
' .Num_Meses_PI = 0
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0088 |
' .Formula = ""
' .COD_CLASSE = 1
' .Val_Fator_Franq = 0
' .Alterar_Cobert = ""
' .Val_Risco_Cob = 2000
' .Nom_Cobertura = ""
' .Cod_Grupo_Cobertura = 0
' End With
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0089 |
' .ColecaoPedCobert.Add lobjstcPedCobert, lobjstcPedCobert.Class_ID
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0090 |
' Set lobjstcPedCobert = New stcA46V704B
' With lobjstcPedCobert
' .NOSSO_NUMERO = "01129031531009682261"
' .Cod_Cobert = 6
' .Num_Ped = 212
' .Cod_Org = 0
' .Cod_Ramo = 114
' .Num_Apol = 0
' .Num_Endo = 0
' .Num_Item = 1
' .Tab_Cobert = 1
' .Val_IS = 20000
' .Per_Taxa = 0.1467
' .Per_Desc_Coml = 0
' .Per_Desc_Exp = 25
' .VAL_PR = 18.05
' .Val_Pr_Anu = 18.05
' .Val_Pr_Tarif = 18.05
' .Val_Pr_Tarif_Anu = 18.05
' .Val_Pr_Resg = 18.05
' .Val_Pr_Resg_Anu = 18.05
' .Per_Franq = 0
' .Val_Franq_Min = 0
' .Val_Franq_Max = 0
' .NUM_PI = 1100000212
' .Num_Meses_PI = 0
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0091 |
' .Formula = ""
' .COD_CLASSE = 1
' .Val_Fator_Franq = 1
' .Alterar_Cobert = ""
' .Val_Risco_Cob = 20000
' .Nom_Cobertura = ""
' .Cod_Grupo_Cobertura = 0
' End With
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0092 |
' .ColecaoPedCobert.Add lobjstcPedCobert, lobjstcPedCobert.Class_ID
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0093 |
' Set lobjstcPedCobert = New stcA46V704B
' With lobjstcPedCobert
' .NOSSO_NUMERO = "01129031531009682261"
' .Cod_Cobert = 9
' .Num_Ped = 212
' .Cod_Org = 0
' .Cod_Ramo = 114
' .Num_Apol = 0
' .Num_Endo = 0
' .Num_Item = 1
' .Tab_Cobert = 1
' .Val_IS = 10000
' .Per_Taxa = 1
' .Per_Desc_Coml = 0
' .Per_Desc_Exp = 25
' .VAL_PR = 61.54
' .Val_Pr_Anu = 61.54
' .Val_Pr_Tarif = 61.54
' .Val_Pr_Tarif_Anu = 61.54
' .Val_Pr_Resg = 61.54
' .Val_Pr_Resg_Anu = 61.54
' .Per_Franq = 10
' .Val_Franq_Min = 2000
' .Val_Franq_Max = 0
' .NUM_PI = 1100000212
' .Num_Meses_PI = 0
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0094 |
' .Formula = ""
' .COD_CLASSE = 1
' .Val_Fator_Franq = 0
' .Alterar_Cobert = ""
' .Val_Risco_Cob = 10000
' .Nom_Cobertura = ""
' .Cod_Grupo_Cobertura = 0
' End With
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0095 |
' .ColecaoPedCobert.Add lobjstcPedCobert, lobjstcPedCobert.Class_ID
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0096 |
' Set lobjstcPedCobert = New stcA46V704B
' With lobjstcPedCobert
' .NOSSO_NUMERO = "01129031531009682261"
' .Cod_Cobert = 18
' .Num_Ped = 212
' .Cod_Org = 0
' .Cod_Ramo = 114
' .Num_Apol = 0
' .Num_Endo = 0
' .Num_Item = 1
' .Tab_Cobert = 1
' .Val_IS = 10000
' .Per_Taxa = 0.4
' .Per_Desc_Coml = 0
' .Per_Desc_Exp = 25
' .VAL_PR = 24.61
' .Val_Pr_Anu = 24.61
' .Val_Pr_Tarif = 24.61
' .Val_Pr_Tarif_Anu = 24.61
' .Val_Pr_Resg = 24.61
' .Val_Pr_Resg_Anu = 24.61
' .Per_Franq = 10
' .Val_Franq_Min = 400
' .Val_Franq_Max = 0
' .NUM_PI = 1100000212
' .Num_Meses_PI = 0
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0097 |
' .Formula = ""
' .COD_CLASSE = 1
' .Val_Fator_Franq = 0
' .Alterar_Cobert = ""
' .Val_Risco_Cob = 10000
' .Nom_Cobertura = ""
' .Cod_Grupo_Cobertura = 0
' End With
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0098 |
' .ColecaoPedCobert.Add lobjstcPedCobert, lobjstcPedCobert.Class_ID
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0099 |
' Set .ColecaoPedLocBem = New Collection
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0100 |
' Set .ColecaoPedSeguro = New Collection
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0101 |
' Set .ColecaoPedFator = New Collection
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0102 |
' Set .objstcPedObs = New stcA46V714B
' With .objstcPedObs
' .NOSSO_NUMERO = 1.12903153100968E+18
' .Num_Item = 1
' .Observacao = ""
' End With
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0103 |
' Set .objstcPedInsp = New stcA46V717B
' With .objstcPedInsp
' .NOSSO_NUMERO = ""
' .Num_Ped = 0
' .Cod_Org = 0
' .Cod_Ramo = 0
' .Num_Item = 0
' .Num_Insp = 0
' .Dat_Cadastro = 0
' .Pess_Cont = ""
' .Num_Tel = ""
' .Num_Logr_Risco = 0
' End With
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0104 |
' Set .ColecaoPedClau = New Collection
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0105 |
' Set .ColecaoNormas = New Collection
'
''| 17:24:10 | felipe2              | WDC10900FSLU166  | D46O003      | 0106 |
' End With
' .ColecaoPedLoc.Add lobjstcPedLoc, lobjstcPedLoc.Class_ID
' End With
'
'End Function
'
'Private Sub Command2_Click()
'
'Dim objstc  As stcA46V702B
'Dim strBenef As String
'Dim lStcTextoEspec As stcA46V726B
'Dim intLoop As Integer
'Dim objA46O003A  As clsA46O003A
'Dim lobjstcCob  As stcA46V704B
'Dim lobjstcLoc  As stcA46V708B
'
'Dim normas  As New Collection
'Dim alertas As New Collection
'Dim erros   As New Collection
'Dim classe4 As Boolean
'Dim strMsg  As String
'
'
'Set objstc = New stcA46V702B
'Set objA46O003A = New clsA46O003A
'
'objA46O003A.gfAutenticarUsuario "test01"
'
'
''Set objstc = objA46O003A.gfSeguroNovo(112)
'Set objstc = objA46O003A.gfProcuraNossoNumero("01400644120000034017", True)
''If objA46O003A.gfSalvar(objstc) = True Then
''    MsgBox "Dados Salvos com sucesso!", vbInformation
''End If
'Set objstc = objA46O003A.gfCalcular("01400644120000034017", normas, alertas, erros, classe4)
'
'
''Set objstc = objA46O003A.gfProRenYasuda("11768332", 1, True)
'
'
'If objA46O003A.UltimoErro <> "" Then
'    MsgBox "Ultimo erro: " & objA46O003A.UltimoErro, vbInformation
'    End
'End If
'If erros.Count > 0 Then
'    strMsg = ""
'    For intLoop = 1 To erros.Count
'        strMsg = strMsg & erros(intLoop) & vbCrLf
'Next
'    MsgBox erros.Count & " Erro(s)." & vbCrLf & strMsg, vbCritical
'    End
'    End If
'If alertas.Count > 0 Then
'    alertas(1).RespostaQuest = 1
'    strMsg = ""
'    For intLoop = 1 To alertas.Count
'        strMsg = strMsg & alertas(intLoop).Mensagem & vbCrLf
'Next
'    MsgBox alertas.Count & " alertas." & vbCrLf & strMsg, vbInformation
'End If
'
'
'If Not objstc Is Nothing Then
'    MsgBox "Num_ped:            " & objstc.Num_Ped & vbCrLf & _
'           "Val_IOF:            " & Format(objstc.Val_IOF, "standard") & vbCrLf & _
'           "Val_Cust_Emis:       " & Format(objstc.Val_Cust_Emis, "standard") & vbCrLf & _
'           "VAL_PR:           " & Format(objstc.VAL_PR, "standard"), vbInformation
'           End
'End If
'
''MsgBox objstc.getRpt.CorretorCod
''For intLoop = 1 To alertas.Count
''     alertas(intLoop).RespostaQuest = 1
''    Next
''Set lobjstcLoc = objstc.ColecaoPedLoc(1)
''For Each lobjstcCob In lobjstcLoc.ColecaoPedCobert
''    If lobjstcCob.getRpt.temFranquia = True Then
''        Debug.Print lobjstcCob.getRpt.codFranquia & "-" & lobjstcCob.getRpt.nomFranquia & "-" & lobjstcCob.getRpt.descFranquia
''End If
''    Next
'
'
''objstc.Beneficiarios = "Marcos de Freitas Barra" & vbCrLf & "jose" & vbCrLf & "augusto" & vbCrLf & "pedro" & vbCrLf
''objstc.gfPosicaoColecao = "Marcos de Freitas Barra - Marcos de Freitas Barra - Marcos de Freitas Barra - Joao da Silva"
'
''Set lStcTextoEspec = objstc.ColecaoTextoEspec
''For Each lStcTextoEspec In objstc.ColecaoTextoEspec
''    Debug.Print lStcTextoEspec.Espec_lin1
''    Debug.Print lStcTextoEspec.Espec_lin2
''    Debug.Print lStcTextoEspec.Espec_lin3
''    Debug.Print lStcTextoEspec.Espec_lin4
''    Debug.Print lStcTextoEspec.Espec_lin5
''    Debug.Print lStcTextoEspec.Espec_lin6
''    Debug.Print lStcTextoEspec.Espec_lin7
''    Debug.Print lStcTextoEspec.Espec_lin8
''    Debug.Print lStcTextoEspec.Espec_lin9
''    Debug.Print lStcTextoEspec.Espec_lin10
''    Debug.Print lStcTextoEspec.Espec_lin11
''    Debug.Print lStcTextoEspec.Espec_lin12
''    Debug.Print lStcTextoEspec.Espec_lin13
''    Debug.Print lStcTextoEspec.Espec_lin14
''    Debug.Print lStcTextoEspec.Espec_lin15
''    Debug.Print lStcTextoEspec.Espec_lin16
''    Debug.Print lStcTextoEspec.Espec_lin17
''    Debug.Print lStcTextoEspec.Espec_lin18
''    Debug.Print lStcTextoEspec.Espec_lin19
''    Debug.Print lStcTextoEspec.Espec_lin20
''Next
'
End Sub

