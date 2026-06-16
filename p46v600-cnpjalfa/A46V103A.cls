VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA46V103A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mClsM46V103 As clsM46V103

Public Property Get Class_Name() As String
    Class_Name = "A46V103A"
End Property

Public Sub gpVerIteNiv_2(pobjstcItem As stcA46V708B, _
                        pobjstcPed As stcA46V702B, _
                        pValISBas As Double, _
                        panLocRiscoEnable As Boolean, _
                        pobjstcUsuario As stcA46V725B, _
                        tagDynaTrace As String)
                        
    Dim lKeyDynaTrace As Long
    M46V999.tagDynaTrace = tagDynaTrace
    Dim methodName As String
    methodName = "gpVerIteNiv_2"
    lKeyDynaTrace = M46V999.DynaTrace.iniciarMetodoPublico(tagDynaTrace, methodName, Class_Name)

'Procedure: consiste item n�vel 2: consiste todos dados do item.
    Dim lobjstcCobert As stcA46V704B
    Dim lDig As String
    Dim clsA99V001 As clsA99V001
    Dim lMensagem As String
    Dim lAuxiliar As String
    Dim lNumCobert As Integer
    Dim lNumCobertAdicionais As Integer
    Dim lArrayCoberts() As Integer
    Dim lArrayISCoberts() As Double
    Dim i As Integer
    Dim lDescontoComercial As Double
    Dim lComissao As Double
    Dim lPerDescLMI As Double
    Dim lobjstcPedClau As stcA46V703B   'Objeto da classe de estrutura stcA46V704B (TAB_PED_CLAU).
    Dim lbdTab_Tarifa As ADODB.Connection    'Arquivo: P00Tarifa.
    Dim lrsTab_Tarifa As clsYasRecordSet    'Registro : TAB_Tarifa.
    Dim lDesconto As Double
    Dim pulaLimite As Boolean
    Dim lValLMG As Long
    Dim lDescParm As Double
    Dim lretorno As String
    Dim dComissaoTipEmiss As String
    Dim TemVerbaSep As Boolean
    Dim lAuxVar As Double
    Dim lSelect As String
    Dim bValidaDescontoComercial As Boolean
    Dim lDescontoWeb As Double

    '1201070 - Tarifa - 1050 e 1051 - Item 1 - Planos TGP
    'Contru��o Mista ou Inferior - Aceitas somente quando os im�veis estiverem localizados nos Estados do Paran�, Santa Catarina, Rio Grande do Sul ou S�o Paulo.
    If pobjstcItem.Cod_Ramo = e_CodRamo_Residencial And _
       pobjstcItem.Cod_Plano >= 2278 And _
       pobjstcItem.Cod_Plano <= 2297 Then
        If pobjstcItem.Cod_Constr = e_CodContr_Mista Or pobjstcItem.Cod_Constr = M46V999.gfBuscarDePara(e_CodContr_Inferior, 4) Then
            Select Case pobjstcItem.Sig_UF_Risco
            Case "PR", "SC", "RS", "SP"
            Case Else
                Call M46V999.gpMenErro("Plano TGP - Contru��o Mista ou Inferior n�o aceita nessa regi�o.")
            End Select
        End If
    End If

    '''consist�ncia de comum/ fonte com M46V999.Checkout- depois colocar no ver_comum... thais
    If pobjstcPed.Cod_Estip = M46V999.gfPlanoBasico(pobjstcPed.Cod_Ramo) And pobjstcPed.Cod_Ramo = e_CodRamo_Roubos And pobjstcPed.Tip_Pessoa = e_TipPessoa_Juridica Then
        Call M46V999.gpMenErro("FGE0094 - Para este plano, o segurado n�o pode ser pessoa jur�dica.")
    End If

    '1. Tipo de Seguro
    '1000347 - Inclus�o de informa��es da cong�nere em renova��es
    If pobjstcItem.Tip_Emissao = e_TipEmissao_RenCongenere And pobjstcItem.Ren_Con_Seguradora = 0 And M46V999.gCodProduto >= 988 Then
        Call M46V999.gpMenErro("FGE0090 - Tipo de Emiss�o (Renova��o Cong�nere). Favor informar a Seguradora Cong�nere.")
    End If

    '2. Classifica��o
    If Not M46V999.gfPreenchido(pobjstcItem.Cod_Clasf) Or _
       Not M46V999.gfPreenchido(pobjstcItem.Cod_Compl_Clasf) Or _
       Trim(pobjstcItem.Cod_Clasf) = "0" Then
        Call M46V999.gpMenErro("FGE0094 - Item " & pobjstcItem.Num_Item & " - Classifica��o inv�lida (Classifica��o n�o preenchida).")
    Else
        'Ficha 1000308 - Permitir itens com planos diferentes no produto Roubo (c�lculo coletivo)
        If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, lbdTab_Tarifa, lMensagem) = False Then
            Call M46V999.gpMenErro(lMensagem)
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, lMensagem, 2, Class_Name, methodName)
            GoTo SAIDA
        End If
        lSelect = "SELECT * FROM Tbm_Classificacao WITH (NOLOCK)"
        lSelect = lSelect & " WHERE COD_RAMO = " & pobjstcItem.Cod_Ramo
        lSelect = lSelect & " AND TAB_CLASF = " & pobjstcItem.Tab_Clasf
        lSelect = lSelect & " AND COD_CLASF = '" & pobjstcItem.Cod_Clasf & "'"
        lSelect = lSelect & " AND COD_COMPL_CLASF = '" & pobjstcItem.Cod_Compl_Clasf & "'"
        lSelect = lSelect & " AND COD_TARIF = " & pobjstcItem.Cod_Iden_Tarif
        lSelect = lSelect & " AND Cod_Sequencia  <= " & M46V999.gCodSequencia
        If M46V999.gfObtRegistro(lbdTab_Tarifa, lSelect, lrsTab_Tarifa, lMensagem) = False Then
            Call M46V999.gpGraLog(0, lMensagem)
            Call M46V999.gpMenErro(lMensagem)
            Call M46V999.gpFecha2(lbdTab_Tarifa)
            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, lMensagem, 2, Class_Name, methodName)
            GoTo SAIDA
        End If
        If lrsTab_Tarifa.EOF Then
            Call M46V999.gpMenErro("FGE0095 - Item " & pobjstcItem.Num_Item & " - Classifica��o inv�lida!")
        End If
    End If

    '3. Endere�o do segurado.
    If panLocRiscoEnable Then
        If Not M46V999.gfPreenchido(pobjstcItem.Cod_Tip_Logr_Risco) Then
            Call M46V999.gpMenErro("FGE0095 - Item " & pobjstcItem.Num_Item & " - Endere�o de risco inv�lido (tipo do logradouro n�o preenchido).")
        End If
        If Not M46V999.gfPreenchido(pobjstcItem.Nom_Logr_Risco) Then
            Call M46V999.gpMenErro("FGE0096 - Item " & pobjstcItem.Num_Item & " - Endere�o de risco inv�lido (nome do logradouro n�o preenchido).")
        End If
        If UCase(pobjstcItem.Cod_Tip_Logr_Risco) = "CX" Or UCase(pobjstcItem.Cod_Tip_Logr_Risco) = "CP" _
           Or InStr(UCase(pobjstcItem.Nom_Logr_Risco), "CAIXA POSTAL") Then
            Call M46V999.gpMenErro("FGE0117 - Item " & pobjstcItem.Num_Item & " - Endere�o de risco n�o pode ser uma caixa postal.")
        End If
        If Not M46V999.gfPreenchido(pobjstcItem.Dsc_Compl_Risco) And ((Trim(pobjstcItem.Cod_Clasf) = "918" And pobjstcItem.Cod_Compl_Clasf = "01") Or (Trim(pobjstcItem.Cod_Clasf) = "922" And pobjstcItem.Cod_Compl_Clasf = "02")) Then
            Call M46V999.gpMenErro("FGE0097 - Item " & pobjstcItem.Num_Item & " - Complemento do endere�o de risco inv�lido (Complemento do endere�o de risco n�o preenchido).")
        End If
        If pobjstcItem.objstcPedInsp.Num_Logr_Risco = 0 And Not M46V999.gfPreenchido(pobjstcItem.Dsc_Compl_Risco) Then
            Call M46V999.gpMenErro("FGE0097 - Item " & pobjstcItem.Num_Item & " - Endere�o de risco inv�lido (N�mero n�o preenchido).")
        End If
        If Not M46V999.gfPreenchido(pobjstcItem.Nom_Bairro_Risco) Then
            Call M46V999.gpMenErro("FGE0098 - Item " & pobjstcItem.Num_Item & " - Endere�o de risco inv�lido (bairro n�o preenchido).")
        End If
        If Not M46V999.gfPreenchido(pobjstcItem.Nom_Cid_Risco) Then
            Call M46V999.gpMenErro("FGE0099 - Item " & pobjstcItem.Num_Item & " - Endere�o de risco inv�lido (cidade n�o " & _
                                   "preenchida).")
        End If
        If Not M46V999.gfPreenchido(pobjstcItem.Sig_UF_Risco) Then
            Call M46V999.gpMenErro("FGE0100 - Item " & pobjstcItem.Num_Item & " - Endere�o de risco inv�lido (unidade da " & _
                                   "federa��o n�o preenchida).")
        End If
        If Not M46V999.gfPreenchido(pobjstcItem.Num_Cep_Risco) Then
            Call M46V999.gpMenErro("FGE0101 - Item " & pobjstcItem.Num_Item & " - Endere�o de risco inv�lido (CEP n�o " & _
                                   "preenchido).")
        End If
        If val(pobjstcItem.Num_Cep_Risco) = 0 Then
            Call M46V999.gpMenErro("FGE0102 - Item " & pobjstcItem.Num_Item & " - Endere�o de risco inv�lido (CEP " & _
                                   "zerado).")
        End If
        If pobjstcItem.Cod_Plano = 2139 Then    'Plano - Bourbon Shopping - Corretora -0499799
            If Not M46V999.gfPreenchido(pobjstcItem.Dsc_Compl_Risco) Then
                Call M46V999.gpMenErro("FGE0103 - Item " & pobjstcItem.Num_Item & " - Complemento de risco n�o informado (N�mero da Loja).")
            End If
        End If

        If M46V999.gfPreenchido(pobjstcItem.Num_Cep_Risco) Then
            If (M46V999.gCodProduto > M46V999.TarifaAntiga And (M46V999.gTipUsuario = e_Tipo_Usu_Corretor Or M46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda)) Or M46V999.gCodProduto >= 940 Then
                If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, lbdTab_Tarifa, lMensagem) = False Then
                    Call M46V999.gpMenErro(lMensagem)
                    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, lMensagem, 2, Class_Name, methodName)
                    GoTo SAIDA
                End If
                lSelect = "Select Cod_Reg from Tbm_Faixa_Regiao "
                lSelect = lSelect & "  where Cod_Ramo = " & pobjstcItem.Cod_Ramo
                lSelect = lSelect & "    and Cod_Modalidade = " & M46V111.gCod_plano
                lSelect = lSelect & "    and Cep_Ini <= " & pobjstcItem.Num_Cep_Risco
                lSelect = lSelect & "    and Cep_Fim >= " & pobjstcItem.Num_Cep_Risco
                lSelect = lSelect & "    and Cod_Tarif = " & pobjstcItem.Cod_Iden_Tarif
                lSelect = lSelect & "    and Cod_Sequencia  <= " & M46V999.gCodSequencia
                lSelect = lSelect & "    order by Cod_Sequencia desc "
                If M46V999.gfObtRegistro(lbdTab_Tarifa, lSelect, lrsTab_Tarifa, lMensagem) = False Then
                    Call M46V999.gpGraLog(0, lMensagem)
                    Call M46V999.gpMenErro(lMensagem)
                    Call M46V999.gpFecha2(lbdTab_Tarifa)
                    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, lMensagem, 2, Class_Name, methodName)
                    GoTo SAIDA
                End If
                'CEP's = Dever�o estar travados os registros de Piau�, Mato Grosso do Sul, Rond�nia, Roraima e Amap�;
                '280 = Piaui
                '320 = Amap�
                '330 = Rond�nia
                '340 = Roraima
                '370 = Mato Grosso do Sul
                '350 = Acr�
                If lrsTab_Tarifa.EOF Then
                    Call M46V999.gpMenErro("FGE0105 - Regi�o n�o cadastrada. Verificar CEP.")
                Else

                    'Solicita��o n� 0801480 est� liberando Amap� (Tarifa 948)
                    If M46V999.gCodProduto < 948 Then
                        'Solicita��o n�0801203
                        If M46V999.Check(lrsTab_Tarifa("Cod_Reg")) = 320 Then
                            'Solicita��o n�801347      M46V999.Check(lrsTab_Tarifa("Cod_Reg")) = 280 Or _
                             'Solicita��o n�801347      M46V999.Check(lrsTab_Tarifa("Cod_Reg")) = 330 Then
                            If M46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then
                                'Amap� est� liberado para Yasuda
                            Else
                                Call M46V999.gpMenErro("N�o foi poss�vel efetuar o c�lculo. Entrar em contato com a Yasuda Mar�tima Seguros.")
                            End If
                        End If
                    End If

                    If (M46V999.Check(lrsTab_Tarifa("Cod_Reg")) = 340 Or _
                        M46V999.Check(lrsTab_Tarifa("Cod_Reg")) = 350) And _
                        (M46V999.gCodProduto < 994) Then    '1000851 - Tarifa Setembro - Permitir CEP Acre e Roraima
                        If M46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then
                            'Acre e Roraima est�o liberados para Yasuda
                        ElseIf pobjstcPed.objstcCtrlEmis.Cod_Corr = 49301 And M46V999.gCodProduto >= 944 Then
                            Call M46V999.gpMenErro("N�o foi poss�vel efetuar o c�lculo. Entrar em contato com a Yasuda Mar�tima Seguros.")
                        Else
                            Call M46V999.gpMenErro("N�o s�o aceitos riscos localizados neste Estado.")
                        End If
                    End If
                    'Ficha 1200833 Retirar consist�ncia
                    '    If M46V999.Check(lrsTab_Tarifa("Cod_Reg")) = 370 Then
                    '       If M46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then
                    'Mato Grosso est� liberado para Yasuda
                    'Ficha 901562
                    '       ElseIf M46V999.gCod_Corretor = 49301 And Val(pobjstcItem.Cod_Clasf) = 22 And Val(pobjstcItem.Cod_Compl_Clasf) = 13 Then
                    'Liberado para a Honda, na classifica��o 022.13
                    'Ficha 1000810
                    '       ElseIf M46V999.gCod_Corretor = 8001 Or M46V999.gCod_Corretor = 8002 Or M46V999.gCod_Corretor = 8003 Or M46V999.gCod_Corretor = 8004 Or M46V999.gCod_Corretor = 8005 Or _
                            M46V999.gCod_Corretor = 8006 Or M46V999.gCod_Corretor = 8008 Or M46V999.gCod_Corretor = 8011 Or M46V999.gCod_Corretor = 8014 Or M46V999.gCod_Corretor = 8028 Then
                    'Liberado para Corretora Volkswagen
                    '       Else
                    '          Call M46V999.gpMenErro("N�o foi poss�vel efetuar o c�lculo. Entrar em contato com a Yasuda Seguros.")
                    '       End If
                    '    End If
                End If

            End If
        End If
    End If
    lNumCobert = 0
    lNumCobertAdicionais = 0
    TemVerbaSep = False
    Dim lValorTotalLMG As Double
    lValorTotalLMG = 0
    If pobjstcPed.Cod_Estip <= 1010 Or pobjstcPed.Cod_Estip = 2300 Then
        pulaLimite = False
        For Each lobjstcCobert In pobjstcItem.ColecaoPedCobert
            If lobjstcCobert.Val_IS = 0 And pobjstcItem.Cod_DizerItem <> 12 Then    'N�o consiste para endosso
                Call M46V999.gpMenErro("FGE0107 - Valor LMI da cobertura " & lobjstcCobert.Cod_Cobert & " inv�lido (n�o informado).")
            End If
            If lobjstcCobert.Val_IS < 100# And pobjstcItem.Cod_DizerItem <> 12 Then
                Call M46V999.gpMenErro("FGE0112 - Valor LMI da cobertura " & lobjstcCobert.Cod_Cobert & " inv�lido (menor que 100,00).")
            End If
            'Consist�ncia de Valores de IS
            If lobjstcCobert.Cod_Ramo = 112 Then
                If lobjstcCobert.Cod_Cobert = M46V999.gfBuscarDePara(78, 1) Or lobjstcCobert.Cod_Cobert = M46V999.gfBuscarDePara(79, 1) Then
                    If (lobjstcCobert.Val_IS > 1000000# And M46V999.gDesconsiderou = 0 And M46V999.gCodProduto < 1080) Or _
                        (lobjstcCobert.Val_IS > 5000000# And M46V999.gDesconsiderou = 0 And M46V999.gCodProduto >= 1080) Then
                        Call M46V999.gpMenErro("FGE0108 - Valor LMI da cobertura " & lobjstcCobert.Cod_Cobert & " inv�lido (acima do Limite Permitido).")
                    End If
                End If
            End If

            'Soma Valor Total LMG
            If lobjstcCobert.gCod_Grupo_Cobert = "1" Or lobjstcCobert.gCod_Grupo_Cobert = "5" Then
                lValorTotalLMG = lValorTotalLMG + lobjstcCobert.Val_IS
            End If

            If lobjstcCobert.Cod_Ramo = 113 And M46V999.gCodProduto > 949 And _
               pobjstcPed.objstcCtrlEmis.Cod_Corr = 15701 And _
               lobjstcCobert.Cod_Cobert >= M46V999.gfBuscarDePara(19, 1) And lobjstcCobert.Cod_Cobert <= M46V999.gfBuscarDePara(31, 1) Then
                pulaLimite = True
            End If
            lNumCobert = lNumCobert + 1
            If lobjstcCobert.gCod_Grupo_Cobert = 5 Then
                TemVerbaSep = True
            End If
            
            ReDim Preserve lArrayCoberts(lNumCobert)
            ReDim Preserve lArrayISCoberts(lNumCobert)
            lArrayCoberts(lNumCobert) = lobjstcCobert.Cod_Cobert
            lArrayISCoberts(lNumCobert) = lobjstcCobert.Val_IS

            '1100389 - Consistencia de Franquias Calculo Interno
            If M46V999.gTipUsuario = e_Tipo_Usu_Yasuda And M46V999.gTipoAcesso = DIGITACAO And pobjstcPed.LMI_Unico <> 1 Then
                Dim lbdP00MPED As ADODB.Connection
                Dim lRstFranquia As clsYasRecordSet
                If M46V999.gfAbrBasDados("P00MPED", M46V999.gTipP00MPED, M46V999.gEndP00MPED, lbdP00MPED, lMensagem) = False Then
                    Call M46V999.gpGraLog(2, lMensagem)
                    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, lMensagem, 2, Class_Name, methodName)
                    GoTo SAIDA
                End If
                'Verificar se existe registro na tabela (Tab_Ped_Num_Copia)
                lSelect = "Select NOSSO_NUMERO_OLD From Tab_Ped_Num_Copia"
                lSelect = lSelect & " Where NOSSO_NUMERO_NEW = '" & pobjstcPed.nosso_numero & "' "
                If M46V999.gfObtRegistro(lbdP00MPED, lSelect, lRstFranquia, lMensagem) = False Then
                    Call M46V999.gpGraLog(2, lMensagem)
                    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, lMensagem, 2, Class_Name, methodName)
                    GoTo SAIDA
                End If
                'Se encontrou registro
                If Not lRstFranquia.EOF Then
                    lSelect = " Select val_franq_min, per_franq  From Tab_Ped_Cobert "
                    lSelect = lSelect & " Where Nosso_Numero = '" & lRstFranquia("NOSSO_NUMERO_OLD") & "' "
                    lSelect = lSelect & " And Num_Item = " & pobjstcItem.Num_Item
                    lSelect = lSelect & " And COD_COBERT = " & lobjstcCobert.Cod_Cobert
                    If M46V999.gfObtRegistro(lbdP00MPED, lSelect, lRstFranquia, lMensagem) = False Then
                        Call M46V999.gpGraLog(2, lMensagem)
                        Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, lMensagem, 2, Class_Name, methodName)
                        GoTo SAIDA
                    End If
                    'valida franquia e percentual
                    If Not lRstFranquia.EOF Then
                        If lRstFranquia("val_franq_min") <> lobjstcCobert.Val_Franq_Min Or _
                           lRstFranquia("per_franq") <> lobjstcCobert.Per_Franq Then
                            'configura mensagem
                            lAuxiliar = "Item " & pobjstcItem.Num_Item & " Cobertura " & lobjstcCobert.Cod_Cobert & _
                                        ", Percentual ou Valor de Franquia M�nima difere ao informado no C�lculo Interno." & vbCrLf & _
                                        "C�lculo Interno: M�nima = " & Format(lRstFranquia("val_franq_min"), "standard") & " Percentual = " & Format(lRstFranquia("per_franq"), "standard") & "%" & vbCrLf & _
                                        "Emiss�o: M�nima = " & Format(lobjstcCobert.Val_Franq_Min, "standard") & " Percentual = " & Format(lobjstcCobert.Per_Franq, "standard") & "%"
                            'If MsgBox(lAuxiliar & vbCrLf & vbCrLf & "Deseja prosseguir a emiss�o com essa diverg�ncia?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                            '    MsgBox "Altere o valor ou percentual de franquia m�nima conforme informado e calcule novamente!", vbInformation, "Aten��o"
                            '    Call M46V999.gpMenErro("FGE9999 - " & lAuxiliar)
                            '    M46V999.gDesabFormsTransmi = False
                            'End If
                            Call M46V999.gpMenErro(lAuxiliar)
                            lAuxiliar = ""
                        End If
                    End If
                End If
            End If
            'Ajuste de consist�ncia de Cobertura somente de concession�ria - Simony
            If lobjstcCobert.Cod_Ramo = 112 Then
                If InStr("0055;0056;0057;0063;0064;0065;0066;0067;0068;0069;0070;0071;0072;0073;0074;0075;0122;0123;0124;0125;0098;0126;0137;0138;0152;0153;0154", Format(lobjstcCobert.Cod_Cobert, "0000")) > 0 Then
                    If val(pobjstcItem.Cod_Clasf) = M46V999.gfBuscarDePara(22, 2) And val(pobjstcItem.Cod_Compl_Clasf) = M46V999.gfBuscarDePara(22, 2, 13) Then
                        'Ok pode contratar
                    Else
                        If (pobjstcPed.objstcCtrlEmis.Cod_Corr = 358301 Or pobjstcPed.objstcCtrlEmis.Cod_Corr_2 = 358301 Or pobjstcPed.objstcCtrlEmis.Cod_Corr_3 = 358301) And lobjstcCobert.Cod_Cobert = M46V999.gfBuscarDePara(55, 1) Then
                            'Ok liberado cobertura 55 para esse corretor
                        Else
                            'Norma 122331
                            'Call M46V999.gpMenErro("FGE0108 - Cobertura " & lobjstcCobert.Cod_Cobert & " v�lida somente para a classifica��o 022 -13 (CONCESSIONARIA).")
                        End If
                    End If
                End If
            End If
        Next lobjstcCobert

        'rodrigoS
        'Consiste valor do LMG/VR
        If Not (M46V999.gCod_Ramo = e_CodRamo_Condominio And M46V999.gCodProduto >= 1012) Then
            If lValorTotalLMG > pobjstcItem.Val_Vr Then
                'Call M46V999.gpMenErro("FGE0114 - Valor LMG n�o pode ser maior do que o Valor do Risco.")
                Call M46V999.gpMenErro("FGE0114 - Valor LMI n�o pode ser maior do que o Valor do Risco.")
            End If
        End If


        'Ficha 900198
        '        If lNumCobert > 22 Then
        '            Call M46V999.gpMenErro("AGE0052 - Quantidade de coberturas selecionadas acima do limite (22). ")
        '        End If
    Else
        If pValISBas = 0 Then
            Call M46V999.gpMenErro("FGE0107 - Valor LMI da cobertura b�sica inv�lido (n�o informado).")
        End If
    End If
    
    For Each lobjstcCobert In pobjstcItem.ColecaoPedCobert
        If lobjstcCobert.gCod_Grupo_Cobert <> 1 And lobjstcCobert.gCod_Grupo_Cobert <> 5 And lobjstcCobert.gCod_Grupo_Cobert <> 4 Then
            lNumCobertAdicionais = lNumCobertAdicionais + 1
        End If
    Next lobjstcCobert
    
    If lNumCobertAdicionais = 0 Then 'And M46V999.gTipoAcesso = CALCULO Then
        Select Case M46V999.gCod_Ramo
        Case e_CodRamo_Roubos, e_CodRamo_Riscos_diversos
            'liberado
        Case e_CodRamo_Condominio
            If M46V111.gCod_plano = 2 Then
                'liberado
            Else
                Call M46V999.gpMenErro("FGE0107 - Selecionar pelo menos uma cobertura adicional.")
            End If
        Case Else
            If pobjstcPed.Cod_Estip <> 20 _
            And Not (pobjstcPed.Cod_Estip >= 5061 And pobjstcPed.Cod_Estip <= 5064) _
            Then 'N�o validar Imobiliario
                Call M46V999.gpMenErro("FGE0107 - Selecionar pelo menos uma cobertura adicional.")
            End If
        End Select
    End If
    If lNumCobert > 0 And pobjstcPed.StatusCotSiscota <> enum_StatusTransm_EnvioKit Then
        For i = 1 To lNumCobert
            Call mpRelationCobert(lArrayCoberts(i), lArrayISCoberts(i), lArrayCoberts, lArrayISCoberts, lNumCobert)
        Next i
    End If

    If M46V999.gTipUsuario = e_Tipo_Usu_Corretor Or M46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda Then
        'smrosa - Ficha 801594
        If pulaLimite = False Or _
           (pulaLimite = True And M46V999.gCodProduto < 950) Then
            Call mpLimiteLMG(pobjstcPed, pobjstcItem)
        End If
    End If

    'Verifica Limite de Desconto Comercial do Corretor
    If M46V999.gCodProduto > M46V999.TarifaAntiga Then
        Call mpLimiteComercial(lDescontoComercial, lComissao, pobjstcPed.objstcCtrlEmis.Cod_Corr, pobjstcItem.Cod_Clasf, pobjstcItem.Cod_Compl_Clasf, lPerDescLMI)

        ' ABCABCABC - Verificar se h� limite especificado para esta cota��o espec�fica, e ent�o definir o limite
        If pobjstcPed.objstcCtrlEmis.Per_Max_Comis_Corr > 0 Then
            lComissao = pobjstcPed.objstcCtrlEmis.Per_Max_Comis_Corr
        End If

        lDesconto = pobjstcPed.objstcCtrlEmis.Per_Desc_Coml
        If pobjstcItem.ColecaoPedCobert.Count > 0 Then
            Set lobjstcCobert = pobjstcItem.ColecaoPedCobert.Item(1)
            lDesconto = lobjstcCobert.Per_Desc_Coml
            lDescontoWeb = lobjstcCobert.Per_Desc_Coml_Web
        End If
        
        For Each lobjstcCobert In pobjstcItem.ColecaoPedCobert
            If lobjstcCobert.Cod_Cobert = 1 Then
                lValLMG = lobjstcCobert.Val_IS
                Exit For
            End If
        Next lobjstcCobert

        Dim lClassif As String
        'lClassif = Val(Trim(Left(frmT46V101A.txtClassificacao.Text, 3)) + Trim(Mid(frmT46V101A.txtClassificacao.Text, 6, 2)))

        lClassif = pobjstcItem.Cod_Clasf & pobjstcItem.Cod_Compl_Clasf

        'Ficha 1000408


        If pobjstcPed.Cod_Produto >= 998 Then
            '1001106 - Renova��o Roubo
            If pobjstcItem.ColecaoPedCobert.Count > 0 Then
                Set lobjstcCobert = pobjstcItem.ColecaoPedCobert.Item(1)
                Call M46V103.gfVerParametro(10, pobjstcPed, pobjstcItem, lobjstcCobert, lretorno, lMensagem)
                lDescParm = val(lretorno)
            End If
        Else
            Call M46V111.gfBuscaDescontoParm(M46V999.gCod_Corretor, pobjstcPed.Cod_Unid_Prod, lDescParm, lValLMG, lClassif, pobjstcPed.Cod_Segurado, lMensagem)
        End If
        If lDescParm > 0 And M46V999.gCodProduto < 1080 Then
            lDescontoComercial = lDescParm
            If (lDesconto > lDescParm) And (Not M46O001.PermiteQQDescontoComercial) Then    '1100931 - Permitir Desconto Versao Interna
                Call M46V999.gpMenErro("FGE0116 - Desconto Comercial inv�lido para esta rubrica/LMG ou Segurado" & lMensagem)
            End If
        End If


        'smrosa
        'Ficha 900451
        If M46V999.gCod_Ramo = e_CodRamo_Residencial Then
            If pobjstcItem.Tip_Emissao = e_TipEmissao_RenYasuda Then
                If M46V999.gTipUsuario = e_Tipo_Usu_Corretor Or M46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda Then
                    If pobjstcPed.objstcCtrlEmis.Cod_Corr = 18301 Then
                        If Trim(pobjstcItem.Cod_Clasf) = M46V999.gfBuscarDePara("001", 2) And Trim(pobjstcItem.Cod_Compl_Clasf) = M46V999.gfBuscarDePara("001", 2, "12") Then
                            If lDescontoComercial < 50 Then
                                lDescontoComercial = 50     'se for contaget, renova��o com casa eventual, o limite � 50.
                            End If
                        End If
                    End If
                End If
            End If
        End If

        'Ficha 901772
        If M46V999.gCod_Ramo = e_CodRamo_Residencial Then
            If M46V999.gTipUsuario = e_Tipo_Usu_Corretor Or M46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda Then
                If pobjstcPed.objstcCtrlEmis.Cod_Corr = 1601 Then
                    If (Trim(pobjstcItem.Cod_Clasf) = M46V999.gfBuscarDePara("001", 2) And Trim(pobjstcItem.Cod_Compl_Clasf) = M46V999.gfBuscarDePara("001", 2, "11")) Or _
                       (Trim(pobjstcItem.Cod_Clasf) = "001" And Trim(pobjstcItem.Cod_Compl_Clasf) = "21") Then
                        If lDescontoComercial < 40 Then
                            lDescontoComercial = 40     'se for Seguratec, casa habitual, o limite � 40.
                        End If
                    End If
                End If
            End If
        End If

        '20170620 - Retorno de classe 4, n�o validar o limite de desconto comercial. Apenas manter o limite definido pelo subscritor
        bValidaDescontoComercial = True
        If M46V999.gEhRetornoSiscota(pobjstcPed.nosso_numero) Then
            bValidaDescontoComercial = False
        End If 'Fim retorno desconto comercial classe 4
       
        If bValidaDescontoComercial Then
            'Ficha 1000313
            If M46V999.gCod_Ramo = e_CodRamo_Residencial Or M46V999.gCod_Ramo = e_CodRamo_Empresarial And M46V999.gCodProduto < 1080 Then
                If M46V999.gTipUsuario = e_Tipo_Usu_Corretor Or M46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda Then
                    If pobjstcPed.objstcCtrlEmis.Cod_Corr = 425888 Then    'And M46V111.gCod_Plano <> 2168 And M46V111.gCod_Plano <> 2169 Then
                        Dim lDesconto_Albertoni As Double
                        If lDesconto > lDescontoComercial Then
                            Call M46V999.gpMenErro("FGE0116 - Desconto Comercial inv�lido (acima do Limite Permitido).")
                            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, "FGE0116 - Desconto Comercial inv�lido (acima do Limite Permitido).", 2, Class_Name, methodName)
                            GoTo SAIDA
                        End If
                        lComissao = (pobjstcPed.objstcCtrlEmis.Per_Corr + _
                                     pobjstcPed.objstcCtrlEmis.Per_Corr_2 + _
                                     pobjstcPed.objstcCtrlEmis.Per_Corr_3)
                        If lComissao <= 40 Then lDesconto_Albertoni = 15
                        If lComissao > 40 Then lDesconto_Albertoni = 0
                        If lDesconto > lDesconto_Albertoni Then
                            Call M46V999.gpMenErro("FGE0116 - Desconto Comercial inv�lido (acima do Limite Permitido).")
                            Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, False, lMensagem, 2, Class_Name, methodName)
                            GoTo SAIDA
                        End If
                        GoTo Ver_Comissao
                    End If
                End If
            End If
    
            If lDescontoComercial < lDesconto And M46V999.gCodProduto < 1080 Then
                If Not M46O001.PermiteQQDescontoComercial Then
                    If pobjstcPed.SiscMaxDesc = 0 Then
                        '1101015 - Projeto 11010 - Renova��o Expressa
                        If pobjstcItem.Cod_Iden_Tarif >= 1026 Then
                            If (pobjstcItem.Tip_Emissao <> e_TipEmissao_Endosso) And _
                               ((pobjstcItem.Tip_Emissao <> e_TipEmissao_RenYasuda) Or _
                               (pobjstcItem.Apol_Ant_IND_SINISTRO = 1) Or _
                               (lDesconto > pobjstcItem.Apol_Ant_PER_DESC_COML)) Then
                                Call M46V999.gpMenErro("FGE0116 - Desconto Comercial inv�lido (acima do Limite Permitido).")
                            End If
                        Else
                            Call M46V999.gpMenErro("FGE0116 - Desconto Comercial inv�lido (acima do Limite Permitido).")
                        End If
                    ElseIf pobjstcPed.SiscMaxDesc < lDesconto Then
                        Call M46V999.gpMenErro("FGE0116 - Desconto m�ximo liberado pela �rea t�cnica = " & pobjstcPed.SiscMaxDesc & "%")
                    End If
                End If
            End If
    
            'FRH
            If lDesconto < 0 Then
                Call M46V999.gpMenErro("FGE0116 - Desconto Comercial inv�lido (Valor n�o permitido).")
            End If
            
            'Vila Velha nova tarifa 1094 limitar desconto comercial em 25%
            If UCase(App.EXEName) Like "D46O003*" And lDescontoComercial < lDesconto And pobjstcItem.Cod_Ramo = 114 And pobjstcItem.Cod_Iden_Tarif >= 1094 And pobjstcPed.objstcCtrlEmis.Cod_Corr = 682987 Then
                Call M46V999.gpMenErro("FGE0116 - Desconto Comercial inv�lido (acima do Limite Permitido).")
            End If
            
            'Tarifa 11.12 BOAS - Limitar desconto comercial em 10%
            If UCase(App.EXEName) Like "D46O003*" And lDescontoComercial < lDesconto And pobjstcItem.Cod_Iden_Tarif >= 1112 And pobjstcItem.Cod_Plano = 5003 And pobjstcPed.objstcCtrlEmis.Cod_Corr = 930337 Then
                Call M46V999.gpMenErro("FGE0116 - Desconto Comercial inv�lido (acima do Limite Permitido).")
            End If
            
            'Tarifa 11.20 3. Ajuste Pontual Condominio: Desconto Comercial em 15%
            If UCase(App.EXEName) Like "D46O003*" And lDescontoComercial < lDesconto And pobjstcItem.Cod_Ramo = 114 And pobjstcItem.Cod_Iden_Tarif >= 1120 And (pobjstcPed.objstcCtrlEmis.Cod_Corr = 908212 Or pobjstcPed.objstcCtrlEmis.Cod_Corr = 32201 Or pobjstcPed.objstcCtrlEmis.Cod_Corr = 918390) Then
                Call M46V999.gpMenErro("FGE0116 - Desconto Comercial inv�lido (acima do Limite Permitido).")
            End If
            
            'Tarifa 11.22 12. Incluir desconto Comercial para Corretor Andrea Hilbk em 15%
            If UCase(App.EXEName) Like "D46O003*" And lDescontoComercial < lDesconto And pobjstcItem.Cod_Ramo = 114 And pobjstcItem.Cod_Iden_Tarif >= 1122 And pobjstcPed.objstcCtrlEmis.Cod_Corr = 915926 Then
                Call M46V999.gpMenErro("FGE0116 - Desconto Comercial inv�lido (acima do Limite Permitido).")
            End If
            
            'Tarifa 11.12 2.19    Ajuste Pontual Condom�nio - Concess�o de Desconto Comercial - Corretor ENERGY BROKER
            If UCase(App.EXEName) Like "D46O003*" And lDescontoComercial < lDesconto And pobjstcItem.Cod_Iden_Tarif >= 1112 And pobjstcItem.Cod_Plano < 1000 And pobjstcPed.objstcCtrlEmis.Cod_Corr = 882599 Then
                Call M46V999.gpMenErro("FGE0116 - Desconto Comercial inv�lido (acima do Limite Permitido).")
            End If
            
            'Tarifa 11.26   6 - Inclus�o de desconto comercial para JBR Corretora 15% de desconto
            If UCase(App.EXEName) Like "D46O003*" And lDescontoComercial < lDesconto And pobjstcItem.Cod_Ramo = 114 And pobjstcItem.Cod_Iden_Tarif >= 1126 And pobjstcPed.objstcCtrlEmis.Cod_Corr = 929778 Then
                Call M46V999.gpMenErro("FGE0116 - Desconto Comercial inv�lido (acima do Limite Permitido).")
            End If
            
            'Tarifa 11.30   2 - Inclus�o de desconto comercial para Igrejas - Corretor Monte Libano
            If UCase(App.EXEName) Like "D46O003*" And lDescontoComercial < lDesconto And M46V111.gCod_plano < 20 And pobjstcItem.Cod_Clasf = 216 And pobjstcItem.Cod_Compl_Clasf = "01" And pobjstcItem.Cod_Iden_Tarif >= 1130 And pobjstcPed.objstcCtrlEmis.Cod_Corr = 374120 Then
                Call M46V999.gpMenErro("FGE0116 - Desconto Comercial inv�lido (acima do Limite Permitido).")
            End If
            
            'Tarifa 11.34   8 - Concess�o de desconto comercial de at� 20% para a corretora Triunfo C�d, 0917581 no Produto Condom�nio
            'Tranferencia da regra de 917581 - Triunfo para 939571 - BR Insurance - 1174
            If UCase(App.EXEName) Like "D46O003*" And lDescontoComercial < lDesconto And pobjstcItem.Cod_Ramo = 114 And pobjstcItem.Cod_Iden_Tarif >= 1134 And pobjstcPed.objstcCtrlEmis.Cod_Corr = 939571 Then
                Call M46V999.gpMenErro("FGE0116 - Desconto Comercial inv�lido (acima do Limite Permitido).")
            End If
            
            'Tarifa 11.76 - 10  Inclus�o de Desconto Produto Condom�nio - Corretor 917372 - APSA
            If UCase(App.EXEName) Like "D46O003*" And lDescontoComercial < lDesconto And pobjstcItem.Cod_Ramo = 114 And pobjstcItem.Cod_Iden_Tarif >= 1176 And pobjstcPed.objstcCtrlEmis.Cod_Corr = 917372 Then
                Call M46V999.gpMenErro("FGE0116 - Desconto Comercial inv�lido (acima do Limite Permitido).")
            End If
            
            'Tarifa 1178 - 10 Inclus�o do Desconto Residencial para Corretora 0889199 - PRISMINAS
            'If lDescontoComercial < lDesconto And _
            'pobjstcItem.Cod_Ramo = 113 And _
            'pobjstcItem.Cod_Iden_Tarif >= 1178 And _
            'pobjstcPed.objstcCtrlEmis.Cod_Corr = 889199 And _
            '(M46V111.gCod_plano = 10 Or M46V111.gCod_plano = 30) Then
            '    Call M46V999.gpMenErro("FGE0116 - Desconto Comercial inv�lido (acima do Limite Permitido).")
            'End If
                        
            'Tarifa 11.36 11. Ajuste Pontual Condominio: Desconto Comercial em 25% p/ 32201
            If UCase(App.EXEName) Like "D46O003*" And lDescontoComercial < lDesconto And pobjstcItem.Cod_Ramo = 114 And pobjstcItem.Cod_Iden_Tarif >= 1136 And pobjstcPed.objstcCtrlEmis.Cod_Corr = 32201 Then
                Call M46V999.gpMenErro("FGE0116 - Desconto Comercial inv�lido (acima do Limite Permitido).")
            End If
            
            'Tarifa 11.36 3. Ajuste Pontual Condominio: Desconto Comercial em 15% p/ 908762 e 933107
            If UCase(App.EXEName) Like "D46O003*" And lDescontoComercial < lDesconto And pobjstcItem.Cod_Ramo = 114 And pobjstcItem.Cod_Iden_Tarif >= 1136 And (pobjstcPed.objstcCtrlEmis.Cod_Corr = 908762 Or pobjstcPed.objstcCtrlEmis.Cod_Corr = 933107) Then
                Call M46V999.gpMenErro("FGE0116 - Desconto Comercial inv�lido (acima do Limite Permitido).")
            End If
            
            'Tarifa 11.46 2. Disponibilizar o Desconto comercial de at� 10% para a atividade Igreja 216-01 para a corretora   C�d. 708301 - H&H COR SEG LT
            If UCase(App.EXEName) Like "D46O003*" And lDescontoComercial < lDesconto And pobjstcItem.Cod_Ramo = 112 And _
              (pobjstcItem.ObjPai.Cod_Estip = 10 Or pobjstcItem.ObjPai.Cod_Estip = 11) And pobjstcItem.Cod_Iden_Tarif >= 1146 And _
              pobjstcPed.objstcCtrlEmis.Cod_Corr = 708301 And pobjstcItem.Cod_Clasf = "216" And pobjstcItem.Cod_Compl_Clasf = "01" Then
                Call M46V999.gpMenErro("FGE0116 - Desconto Comercial inv�lido (acima do Limite Permitido).")
            End If
            
            'Tarifa 11.46 6. Liberar um desconto de 15% para a corretora Lazam MDS Cor. Adm Seg S.A - C�digo - 885877
            If Not pobjstcUsuario Is Nothing Then
                If UCase(App.EXEName) Like "D46O003*" And lDescontoComercial < lDesconto And pobjstcItem.Cod_Ramo = 114 And _
                  (pobjstcItem.ObjPai.Cod_Estip = 10 Or pobjstcItem.ObjPai.Cod_Estip = 11) And pobjstcItem.Cod_Iden_Tarif >= 1146 And _
                  pobjstcPed.objstcCtrlEmis.Cod_Corr = 885877 And pobjstcUsuario.login_user_rede = "lmds23" Then
                    Call M46V999.gpMenErro("FGE0116 - Desconto Comercial inv�lido (acima do Limite Permitido).")
                End If
            End If
            
            'Tarifa 11.54  Item 2 - Conceder Desconto Comercial
            If UCase(App.EXEName) Like "D46O003*" And lDescontoComercial < lDesconto And pobjstcItem.Cod_Ramo = 114 And pobjstcItem.Cod_Iden_Tarif >= 1154 And _
               ((pobjstcPed.objstcCtrlEmis.Cod_Corr = 923336 And (pobjstcPed.Cod_Estip = 10 Or pobjstcPed.Cod_Estip = 11)) Or _
               (pobjstcPed.objstcCtrlEmis.Cod_Corr = 921007 And pobjstcPed.Cod_Estip = 11)) Then
                Call M46V999.gpMenErro("FGE0116 - Desconto Comercial inv�lido (acima do Limite Permitido).")
            End If
                            
            'Tarifa 11.58  Item 5 - Conceder Desconto Comercial - "914347","930781","918351","918462"
            If UCase(App.EXEName) Like "D46O003*" And lDescontoComercial < lDesconto And pobjstcItem.Cod_Ramo = 114 And pobjstcItem.Cod_Iden_Tarif >= 1158 And _
               (pobjstcPed.objstcCtrlEmis.Cod_Corr = 914347 Or pobjstcPed.objstcCtrlEmis.Cod_Corr = 930781 Or pobjstcPed.objstcCtrlEmis.Cod_Corr = 918351 Or pobjstcPed.objstcCtrlEmis.Cod_Corr = 918462) And (pobjstcPed.Cod_Estip = 10 Or pobjstcPed.Cod_Estip = 11) Then
                Call M46V999.gpMenErro("FGE0116 - Desconto Comercial inv�lido (acima do Limite Permitido).")
            End If
                            
            'Tarifa 11.36 3. Ajuste Pontual Condominio: Desconto Comercial em 15% p/ 908762 e 933107
            If UCase(App.EXEName) Like "D46O003*" And lDescontoComercial < lDesconto And pobjstcItem.Cod_Ramo = 114 And pobjstcItem.Cod_Iden_Tarif >= 1156 And (pobjstcPed.objstcCtrlEmis.Cod_Corr = 919155) Then
                Call M46V999.gpMenErro("FGE0116 - Desconto Comercial inv�lido (acima do Limite Permitido).")
            End If
            
            'Tarifa 11.74. Ajuste Pontual Condominio: Desconto Comercial em 10% p/ 879099
            If UCase(App.EXEName) Like "D46O003*" And lDescontoComercial < lDesconto And pobjstcItem.Cod_Ramo = 114 And pobjstcItem.Cod_Iden_Tarif >= 1174 And (pobjstcPed.objstcCtrlEmis.Cod_Corr = 879099) Then
                Call M46V999.gpMenErro("FGE0116 - Desconto Comercial inv�lido (acima do Limite Permitido).")
            End If
            
            'Tarifa 11.66 3. Ajuste Pontual Condominio: Desconto Comercial em 15% p/ 923335 - PROTEST COR SEG LT
            If UCase(App.EXEName) Like "D46O003*" And lDescontoComercial < lDesconto And pobjstcItem.Cod_Ramo = 114 And pobjstcItem.Cod_Iden_Tarif >= 1166 And (pobjstcPed.objstcCtrlEmis.Cod_Corr = 923335) Then
                Call M46V999.gpMenErro("FGE0116 - Desconto Comercial inv�lido (acima do Limite Permitido).")
            End If
            
            'Tarifa 11.68 3. Ajuste Pontual Condominio: Desconto Comercial em 15% p/ 921007 - AUXILIADORA COR SEG LT
            If lDescontoComercial < lDesconto And pobjstcItem.Cod_Ramo = 114 And pobjstcItem.Cod_Iden_Tarif >= 1166 And (pobjstcPed.objstcCtrlEmis.Cod_Corr = 921007) Then
                Call M46V999.gpMenErro("FGE0116 - Desconto Comercial inv�lido (acima do Limite Permitido).")
            End If
            
            'Tarifa 11.72 Ajuste Pontual Condominio: Desconto Comercial em 10% p/ 920770,921012 e 920976
            If lDescontoComercial < lDesconto And pobjstcItem.Cod_Ramo = 114 And pobjstcItem.Cod_Iden_Tarif >= 1172 _
                And (pobjstcPed.objstcCtrlEmis.Cod_Corr = 920770 Or _
                pobjstcPed.objstcCtrlEmis.Cod_Corr = 921012 Or _
                pobjstcPed.objstcCtrlEmis.Cod_Corr = 920976) _
            Then
                Call M46V999.gpMenErro("FGE0116 - Desconto Comercial inv�lido (acima do Limite Permitido).")
            End If
            
            'Shelter Limitar desconto comercial em 15% condominio residencial - Tarifa 1104
            If pobjstcItem.Cod_Iden_Tarif >= 1104 And _
               (Trim(pobjstcPed.ColecaoPedLoc(1).Cod_Clasf) <> "801" And _
                Trim(pobjstcPed.ColecaoPedLoc(1).Cod_Clasf) <> "809" And _
                Trim(pobjstcPed.ColecaoPedLoc(1).Cod_Clasf) <> "810" And _
                pobjstcPed.objstcCtrlEmis.Cod_Corr <> 921012 And _
                (pobjstcPed.objstcCtrlEmis.Cod_Corr <> 908178 And pobjstcItem.Cod_Iden_Tarif >= 1132)) Then
                lDescontoComercial = 0
            End If
            
            If lDescontoComercial < lDesconto And pobjstcItem.Cod_Iden_Tarif >= 1104 And pobjstcPed.objstcCtrlEmis.Cod_Corr = 908178 Then
                Call M46V999.gpMenErro("FGE0116 - Desconto Comercial inv�lido (acima do Limite Permitido).")
            End If
    
            'Corretora Sherrer & Sherrer - Tarifa 1106 '8000047472, Corre��o de desconto - Corretor Scherrer
            If (M46V999.gCod_Ramo = e_CodRamo_Condominio) And lDescontoComercial < lDesconto And pobjstcItem.Cod_Iden_Tarif >= 1106 And pobjstcPed.objstcCtrlEmis.Cod_Corr = 921012 Then
                Call M46V999.gpMenErro("FGE0116 - Desconto Comercial inv�lido (acima do Limite Permitido).")
            End If
        End If '#FIm bValidaDescontoComercial
        
Ver_Comissao:
        '1001224 - Condi��es Diferenciada - Seguratec
        Select Case pobjstcPed.Cod_Segurado
        Case 777787, 777733, 777714, 777713, 777712, 777711, 777710, 777706, 777704, 777702, 777701, 777707, 777716, 777788
            If pobjstcPed.Cod_Produto < 1000 Then GoTo ValidarSeguratec
        Case Else
            'Ficha 901772
ValidarSeguratec:
            If M46V999.gCod_Ramo = e_CodRamo_Residencial And M46V111.gCod_plano <> 2157 And M46V111.gCod_plano <> 2264 Then
                If M46V999.gTipUsuario = e_Tipo_Usu_Corretor Or M46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda Then
                    If pobjstcPed.objstcCtrlEmis.Cod_Corr = 1601 Then
                        If (Trim(pobjstcItem.Cod_Clasf) = M46V999.gfBuscarDePara("001", 2) And Trim(pobjstcItem.Cod_Compl_Clasf) = M46V999.gfBuscarDePara("001", 2, "11")) Or _
                           (Trim(pobjstcItem.Cod_Clasf) = M46V999.gfBuscarDePara("001", 2) And Trim(pobjstcItem.Cod_Compl_Clasf) = M46V999.gfBuscarDePara("001", 2, "21")) Then
                            'Ficha 901836
                            Select Case lDesconto
                            Case 0
                                lComissao = 40    'se for Seguratec, casa ou apto habitual, com Desconto de 0%, o limite de Comiss�o � 40%.
                            Case Is > 0
                                If lDesconto <= 30 Then lComissao = 25  'se for Seguratec, casa ou apto habitual, com Desconto at� 30%, o limite de Comiss�o � 25%.
                                If lDesconto > 30 And lDesconto <= 40 Then lComissao = 15    'se for Seguratec, casa ou apto habitual, com Desconto de 30% a 40%, o limite de Comiss�o � 15%.
                                If lDesconto > 40 Then lComissao = 0    'se for Seguratec, casa ou apto habitual, com Desconto acima de 40%, o limite de Comiss�o � 0%.
                            End Select
                        End If
                    End If
                End If
            End If
        End Select

        'CR3939 - Comissao Especial - Inicio
        dComissaoTipEmiss = -1
        If M46V103.mfAjustesPontuais(e_Tip_Valor_Comissao_TipEmis, TpoValor, dComissaoTipEmiss, lMensagem, pobjstcItem, 0, tagDynaTrace) = False Then
            Call M46V999.gpGraLog(2, lMensagem)
            'Exit Function
        End If

        'CR3939 - Fim

        'Ficha 1000313
        If M46V999.gCod_Ramo = e_CodRamo_Residencial Or M46V999.gCod_Ramo = e_CodRamo_Empresarial Then
            If M46V999.gTipUsuario = e_Tipo_Usu_Corretor Or M46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda Then
                'Corre��o do IF para Plano da Albertoni
                If pobjstcPed.objstcCtrlEmis.Cod_Corr = 425888 Then  'And M46V111.gCod_Plano <> 2168 And M46V111.gCod_Plano <> 2169 Then
                    Dim lComissao_Albertoni As Double
                    Select Case lDesconto
                    Case 0
                        lComissao_Albertoni = 60    'se for Albertoni, com Desconto de 0%, o limite de Comiss�o � 60%.
                    Case Is > 0
                        If lDesconto <= 15 Then lComissao_Albertoni = 40  'se for Albertoni, Com Desconto at� 15%, o limite de Comiss�o � 40%.
                        If lDesconto > 15 Then lComissao_Albertoni = 0    'se for Albertoni, com Desconto acima de 15%, o limite de Comiss�o � 0%.
                    End Select
                    If lComissao > lComissao_Albertoni Then
                        Call M46V999.gpMenErro("FGE0113 - Comiss�o inv�lida (acima do Limite Permitido).")
                    End If
                    GoTo Ver_LMI
                End If
            End If
        End If

        '1200806 - Tarifa Setembro - 1044 e 1045 - Item 9 - Plano Funcionario Yasuda
        If pobjstcItem.Cod_Iden_Tarif >= 1044 And _
           pobjstcItem.Cod_Ramo = e_CodRamo_Residencial And _
           pobjstcPed.objstcCtrlEmis.Cod_Corr = 1601 Then
            Select Case pobjstcPed.Cod_Segurado
            Case 954815, 954824, 954816, 954817
                'valida comiss�o
                If (pobjstcPed.objstcCtrlEmis.Per_Corr + pobjstcPed.objstcCtrlEmis.Per_Corr_2 + pobjstcPed.objstcCtrlEmis.Per_Corr_3) <> 10 Then
                    Call M46V999.gpMenErro("FGE0113 - Comiss�o inv�lida, fora das condi��es para esse segurado (Comiss�o = 10% e Desconto 40%).")
                End If

                'N�o validar desconto retorno classe 4
                If bValidaDescontoComercial Then
                    If lDesconto <> 40 Then
                        Call M46V999.gpMenErro("FGE0115 - Desconto Comercial inv�lido, fora das condi��es para esse segurado (Comiss�o = 10% e Desconto 40%)." & lMensagem)
                    End If
                End If
            End Select
        End If

        'Tarifa 11.12 Habilitar Co-Corretagem no Plano Harmonia
        If pobjstcItem.Cod_Ramo = e_CodRamo_Residencial And lComissao > (pobjstcPed.objstcCtrlEmis.Per_Corr + pobjstcPed.objstcCtrlEmis.Per_Corr_2 + pobjstcPed.objstcCtrlEmis.Per_Corr_3) And pobjstcItem.Cod_Iden_Tarif >= 1112 And _
           (pobjstcItem.Cod_Plano >= 2476 And pobjstcItem.Cod_Plano <= 2483) Then
            Call M46V999.gpMenErro("FGE0113 - Comiss�o inv�lida (abaixo do valor estabelecido).")
        End If

        'Tarifa 11.14 Minha casa minha vida
        If pobjstcItem.Cod_Ramo = e_CodRamo_Condominio And lComissao > (pobjstcPed.objstcCtrlEmis.Per_Corr + pobjstcPed.objstcCtrlEmis.Per_Corr_2 + pobjstcPed.objstcCtrlEmis.Per_Corr_3) And pobjstcItem.Cod_Iden_Tarif >= 1114 And _
           (pobjstcItem.Cod_Plano >= 2051 And pobjstcItem.Cod_Plano <= 2057) And pobjstcPed.objstcCtrlEmis.Cod_Corr = "0852557" Then
            Call M46V999.gpMenErro("FGE0113 - Comiss�o inv�lida (abaixo do valor estabelecido).")
        End If
        
        If lComissao < (pobjstcPed.objstcCtrlEmis.Per_Corr + _
                        pobjstcPed.objstcCtrlEmis.Per_Corr_2 + _
                        pobjstcPed.objstcCtrlEmis.Per_Corr_3) Then
            '1101015 - Projeto 11010 - Renova��o Expressa
            If (pobjstcItem.Cod_Iden_Tarif >= 1026 And pobjstcItem.Tip_Emissao <> e_TipEmissao_SeguroNovo) Then
                If (pobjstcItem.Tip_Emissao <> e_TipEmissao_Endosso) And _
                   ((pobjstcItem.Tip_Emissao <> e_TipEmissao_RenYasuda) Or _
                   (pobjstcItem.Apol_Ant_IND_SINISTRO = 1) Or _
                   (lComissao > pobjstcItem.ObjPai.Apol_Ant_PER_CORR And pobjstcItem.ObjPai.Apol_Ant_PER_CORR > 0) Or _
                   ((pobjstcPed.objstcCtrlEmis.Per_Corr + _
                     pobjstcPed.objstcCtrlEmis.Per_Corr_2 + _
                     pobjstcPed.objstcCtrlEmis.Per_Corr_3) > pobjstcItem.ObjPai.Apol_Ant_PER_CORR And pobjstcItem.ObjPai.Apol_Ant_PER_CORR > 0)) Then
                    If dComissaoTipEmiss <> (pobjstcPed.objstcCtrlEmis.Per_Corr + pobjstcPed.objstcCtrlEmis.Per_Corr_2 + pobjstcPed.objstcCtrlEmis.Per_Corr_3) Then
                        Call M46V999.gpMenErro("FGE0113 - Comiss�o inv�lida (acima do Limite Permitido).")
                    End If
                End If
            Else
                If dComissaoTipEmiss <> (pobjstcPed.objstcCtrlEmis.Per_Corr + pobjstcPed.objstcCtrlEmis.Per_Corr_2 + pobjstcPed.objstcCtrlEmis.Per_Corr_3) Then
                    Call M46V999.gpMenErro("FGE0113 - Comiss�o inv�lida (acima do Limite Permitido).")
                End If
            End If
        End If

        lComissao = (pobjstcPed.objstcCtrlEmis.Per_Corr + _
                     pobjstcPed.objstcCtrlEmis.Per_Corr_2 + _
                     pobjstcPed.objstcCtrlEmis.Per_Corr_3)

        'Retirado pois com o novo produto maritima yasuda v�o poder ter comiss�o com menos de 1%
'        If lComissao < 1 And M46V999.gCodProduto >= 940 And (M46V999.gTipUsuario = e_Tipo_Usu_Corretor Or M46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda) Then
'            Call M46V999.gpMenErro("FGE0114 - Comiss�o m�nima permitida 1%.")
'        End If
        
        'Recolocado a pedido da Luana (01/02/2017)
        If lComissao = 0 Then
            Call M46V999.gpMenErro("FGE0114 - N�o � poss�vel ter comiss�o zero.")
        End If

Ver_LMI:
        If M46V999.gNovaPropostaLMI Then
            If pobjstcPed.Per_Desc_LMI > lPerDescLMI And (M46V999.gTipUsuario = e_Tipo_Usu_Corretor Or M46V999.gTipUsuario = e_Tipo_Usu_CorretorYasuda) Then
                Call M46V999.gpMenErro("FGE0115 - Desconto LMI maior do que o permitido.")
            End If
            If pobjstcPed.Per_Desc_LMI > 99 Then
                Call M46V999.gpMenErro("FGE0116 - Desconto LMI maior do que o permitido.")
            End If
        End If

    End If

    'Ficha 9001320
    M46V111.gDesc_Coml = lDescontoComercial

    If M46V999.gCodProduto <= M46V999.TarifaAntiga Then
        For Each lobjstcPedClau In pobjstcItem.ColecaoPedClau
            'If M46V999.gfPreenchido(pObjstcItem.objstcPedClau.NOSSO_NUMERO) Then
            If lobjstcPedClau.Cod_Clau <> 0 Then
                If Left(lobjstcPedClau.Descricao, 14) = String(14, "0") Then
                    Call M46V999.gpMenErro("FGE0111 - Dados do benefici�rio inv�lidos (CPF/CNPJ n�o preenchido).")
                Else
                    Set clsA99V001 = New clsA99V001
                    Set clsA99V001.M46V999 = M46V999

                    If Left(lobjstcPedClau.Descricao, 3) <> "000" Then
                        Call M46V999.gpForCNPJ(1, Left(lobjstcPedClau.Descricao, 14), lAuxiliar)
                        If Not clsA99V001.DIG_CNPJ(lAuxiliar, _
                                                   lDig, lMensagem) Then
                            Call M46V999.gpMenErro("FGE0111 - Dados do benefici�rio inv�lidos (CNPJ inv�lido).")
                        End If
                    Else
                        Call M46V999.gpForCPF(1, Mid(lobjstcPedClau.Descricao, 4, 11), lAuxiliar)
                        If Not clsA99V001.DIG_CPF(lAuxiliar, _
                                                  lDig, lMensagem) Then
                            Call M46V999.gpMenErro("FGE0111 - Dados do benefici�rio inv�lidos (CPF inv�lido).")
                        End If
                    End If
                End If
                If M46V999.gfPreenchido(Mid(lobjstcPedClau.Descricao, 15)) = False Then
                    Call M46V999.gpMenErro("FGE0111 - Dados do benefici�rio inv�lidos (Nome n�o preenchido).")
                End If
            End If
            'End If
            Exit For
        Next
    End If
    'rodrigoS Merge 14/05/2014
    If pobjstcItem.Cod_Ramo <> e_CodRamo_Roubos Then
        Call mpConsisteBens(1, pobjstcItem)
        If M46V999.gCodProduto >= M46V999.TarifaMigracaoMaritima Then
            Call mpConsistePessoa(pobjstcItem)

            lComissao = (pobjstcPed.objstcCtrlEmis.Per_Corr + _
                         pobjstcPed.objstcCtrlEmis.Per_Corr_2 + _
                         pobjstcPed.objstcCtrlEmis.Per_Corr_3)

            Call mpConsisteProlabore(pobjstcPed.ColecaoProlabore, lComissao, pobjstcPed.Cod_Tip_Doc)

            
            Call mpConsisteLimiteAgravo(pobjstcPed.objstcCtrlEmis)
            
        End If
    End If

    'CONSISTENCIA DOS FATORES DE PROTE��O PARA A TARIFA 1080 PASSOU PARA A WEB
    If M46V999.gCodProduto < 1080 Then
    Call mpConsisteFatoresProtecao(pobjstcItem)
    End If
    '  Call mpConsisteValoremRisco
    Call mpVerificarInspecao(pobjstcItem, pobjstcPed)

    'lucros cessantes
    'rodrigoS
    If pobjstcPed.Cod_Ramo = e_CodRamo_Empresarial And pobjstcPed.Cod_Produto >= 1076 And pobjstcItem.Cod_DizerItem = 0 Then
        Call mpConsisteLucrosCessantes(pobjstcItem, pobjstcPed)
    End If
    
    If pobjstcPed.Cod_Ramo = e_CodRamo_Condominio And pobjstcItem.Cod_Plano = 10 And pobjstcPed.Cod_Produto >= 1080 Then
        Call mpConsisteCondominioDanosEletricos(pobjstcItem)
    End If
    
    If pobjstcPed.Cod_Produto >= 1106 And (pobjstcPed.Cod_Tip_Doc = e_TipDoc_Proposta Or M46V999.gPrograma <> "D046O003") Then
        Dim lClasse4 As Boolean
        
        lClasse4 = pobjstcPed.mStatusDivSiscota = enum_StatusTransm_Solicitada
        
        If Not lClasse4 And (pobjstcItem.Cod_DizerItem <= e_CodDizerInclusao Or M46V999.gTarifa_Apolice >= 1106) Then
            Call mpConsisteRelacaoBicicletaNaoInformada(pobjstcItem)
        End If
        
        Call mpConsisteCoberturaBicicletaNaoInformada(pobjstcItem)
        Call mpConsisteValorBicicleta(pobjstcItem)
        Call mpConsisteDescricaoBicicleta(pobjstcItem)
                
        If pobjstcPed.Cod_Ramo = 113 And pobjstcPed.Cod_Estip = 30 Then
            If Not lClasse4 And (pobjstcItem.Cod_DizerItem <= e_CodDizerInclusao Or M46V999.gTarifa_Apolice >= 1106) Then
                Call mpConsisteRelacaoAllRisksNaoInformada(pobjstcItem)
                Call mpConsisteRelacaoObrasArteNaoInformada(pobjstcItem)
                
                '8000046258, Somat�ria da rela��o de bens - All Risks
                Call mpConsisteSomaRelacaoAllRisks(pobjstcItem)
            End If
            
            Call mpConsisteCoberturaAllRisksNaoInformada(pobjstcItem)
            Call mpConsisteValorAllRisks(pobjstcItem)
            Call mpConsisteDescricaoAllRisks(pobjstcItem)
            
            Call mpConsisteCoberturaObrasArteNaoInformada(pobjstcItem)
            Call mpConsisteValorObrasArte(pobjstcItem)
            Call mpConsisteDescricaoObrasArte(pobjstcItem)
        End If
    End If

    'Call mpConsisteCapitalGlobal

    Call mpVerificarInspecao(pobjstcItem, pobjstcPed)

    '    Dim TemEspec As Boolean
    '    TemEspec = False
    '
    '    'M46V999.gCod_Ramo = e_CodRamo_Roubos 'foi retirada esta obrigatoriedade pois ser� somente cota��o
    '
    '    If M46V999.gCod_Ramo = e_CodRamo_Incendio Or _
         '        M46V999.gCod_Ramo = e_CodRamo_Tumultos Or _
         '        M46V999.gCod_Ramo = e_CodRamo_Riscos_de_Petroleo Or _
         '        M46V999.gCod_Ramo = e_CodRamo_Aeronauticos Or _
         '        M46V999.gCod_Ramo = e_CodRamo_LucrosCessantes Or _
         '        M46V999.gCod_Ramo = e_CodRamo_Fianca_Locataria Or _
         '        M46V999.gCod_Ramo = e_CodRamo_Resp_Civil Or _
         '        M46V999.gCod_Ramo = e_CodRamo_Penhor_rural Or _
         '        M46V999.gCod_Ramo = e_CodRamo_Riscos_de_engenharia Or _
         '        M46V999.gCod_Ramo = e_CodRamo_Riscos_diversos Or _
         '        M46V999.gCod_Ramo = e_CodRamo_Garantia_obrigacoes Then
    '        If Not M46V999.gfDisponivel("frmT46V143B") Then
    '            Call M46V999.gpMenErro("FES0001 - Favor Preencher os dados de Especifica��o.")
    '        Else
    '            For I = 1 To frmT46V143B.CkCodTexto.Count
    '                If frmT46V143B.CkCodTexto(I).Value Then
    '                    TemEspec = True
    '                    Exit For
    '                End If
    '            Next
    '            If Not TemEspec Then
    '                Call M46V999.gpMenErro("FES0002 - Favor Preencher os dados de Especifica��o.")
    '            End If
    '        End If
    '    End If

    'Consiste Comiss�o do Corretor
    If Not M46V103.mfConsisteComissao(pobjstcItem, pobjstcPed, tagDynaTrace) Then
        Call M46V999.gpMenErro("FGE0110 - Comiss�o Inv�lida!")
    End If

    ' Ficha 1300395 - In�cio
    Dim CodCobert As Integer
    Dim codCorretor As Long
    
    If M46V999.gCod_Ramo = e_CodRamo_Empresarial And pobjstcItem.Cod_Iden_Tarif < 1080 And pobjstcItem.Cod_Plano < 2000 Then
        codCorretor = pobjstcPed.objstcCtrlEmis.Cod_Corr
        
        If Not (codCorretor = 358301 Or codCorretor = 358302 Or codCorretor = 358303 Or codCorretor = 358305 _
           Or codCorretor = 358306 Or codCorretor = 358307 Or codCorretor = 358308 Or codCorretor = 700302 _
           Or codCorretor = 711013 Or codCorretor = 711799 Or codCorretor = 715427 Or codCorretor = 724007 _
           Or codCorretor = 737699 Or codCorretor = 24301) Then
            
            If Not (Trim(pobjstcItem.Cod_Clasf) = "022" And (Trim(pobjstcItem.Cod_Compl_Clasf) = "13" Or Trim(pobjstcItem.Cod_Compl_Clasf) = "11")) Then
                
                For Each lobjstcCobert In pobjstcItem.ColecaoPedCobert
                    CodCobert = lobjstcCobert.Cod_Cobert
                If CodCobert = M46V999.gfBuscarDePara(55, 1) Or CodCobert = M46V999.gfBuscarDePara(56, 1) Or CodCobert = M46V999.gfBuscarDePara(57, 1) Or _
                    CodCobert = M46V999.gfBuscarDePara(63, 1) Or CodCobert = M46V999.gfBuscarDePara(64, 1) Or CodCobert = M46V999.gfBuscarDePara(65, 1) Or _
                    CodCobert = M46V999.gfBuscarDePara(66, 1) Or CodCobert = M46V999.gfBuscarDePara(67, 1) Or CodCobert = M46V999.gfBuscarDePara(68, 1) Or _
                    CodCobert = M46V999.gfBuscarDePara(69, 1) Or CodCobert = M46V999.gfBuscarDePara(70, 1) Or CodCobert = M46V999.gfBuscarDePara(71, 1) Or _
                    CodCobert = M46V999.gfBuscarDePara(72, 1) Or CodCobert = M46V999.gfBuscarDePara(73, 1) Or CodCobert = M46V999.gfBuscarDePara(74, 1) Or _
                    CodCobert = M46V999.gfBuscarDePara(75, 1) Or CodCobert = M46V999.gfBuscarDePara(122, 1) Or CodCobert = M46V999.gfBuscarDePara(123, 1) Or _
                    CodCobert = M46V999.gfBuscarDePara(124, 1) Or CodCobert = M46V999.gfBuscarDePara(125, 1) Or CodCobert = M46V999.gfBuscarDePara(98, 1) Or _
                    CodCobert = M46V999.gfBuscarDePara(126, 1) Or CodCobert = M46V999.gfBuscarDePara(127, 1) Or CodCobert = M46V999.gfBuscarDePara(137, 1) Or _
                    CodCobert = M46V999.gfBuscarDePara(138, 1) Or CodCobert = M46V999.gfBuscarDePara(152, 1) Or CodCobert = M46V999.gfBuscarDePara(153, 1) Or _
                    CodCobert = M46V999.gfBuscarDePara(154, 1) Then
            
                        ' Chegou at� aqui, ent�o tem que dar erro
                        Call M46V999.gpMenErro("Norma 122331: A cobertura " & CodCobert & " s� pode ser contratadas para as classifica��es 022-11 e 022-13.")
                    End If
                Next
                
            End If
        End If
    End If

    ' Ficha 1300395 - Fim

    Call mpConsisteLMI(pobjstcItem)

    Call mpConsisteEndosso(pobjstcItem)

    Call M46V999.DynaTrace.finalizarMetodoPublico(lKeyDynaTrace, True, , , Class_Name, methodName)
SAIDA:
    GoSub FechaObj
    Exit Sub
FechaObj:
    Set lobjstcCobert = Nothing
    Set clsA99V001 = Nothing
    Set lobjstcPedClau = Nothing
    Set lbdTab_Tarifa = Nothing
    Set lrsTab_Tarifa = Nothing
    Set lbdP00MPED = Nothing
    Set lRstFranquia = Nothing
    Return
End Sub

'8000046258, Somat�ria da rela��o de bens - All Risks
Private Sub mpConsisteSomaRelacaoAllRisks(pobjstcItem As stcA46V708B)
    Dim Item As stcA46V704B
    Dim iAllRisks As stcA46V747B

    If pobjstcItem.ColecaoAllRisks.Count > 0 Then
        For Each Item In pobjstcItem.ColecaoPedCobert
            If Item.Cod_Cobert = 1172 Then
                
                Dim iTotal As Double
                iTotal = 0
                For Each iAllRisks In pobjstcItem.ColecaoAllRisks
                    iTotal = iTotal + iAllRisks.Valor
                Next
                
                If iTotal <> Item.Val_IS Then
                    Call M46V999.gpMenErro("Consist�ncia do comum - A somat�ria da rela��o de bens de All Risks (R$" & CStr(Format(iTotal, "standard")) & ") � diferente da IS contratada (R$" & Format(Item.Val_IS, "standard") & ").")
                End If
            End If
        Next
    End If

    Set Item = Nothing
End Sub

'******************************ADRIANO************************
Private Sub mpConsisteLMI(pobjstcItem As stcA46V708B)
    Dim lobjstcLMI As stcA46V728B


    'Verificar as informa��es digitadas.
    If pobjstcItem.ColecaoLMI.Count < 2 And M46V999.gTipUsuario = e_Tipo_Usu_Yasuda Then
        'Verifico se � obrigat�rio o preenchimentos dos endere�os
        If pobjstcItem.Num_Cep_Risco = 70000000 And _
           pobjstcItem.Nom_Logr_Risco = "CONFORME ESPECIFICACAO EM ANEXO" And _
           pobjstcItem.Nom_Cid_Risco = "BRASILIA" Then

            Call M46V999.gpMenErro("AGE0082 - Qtde de endere�os de LMI inferior a 2")

        End If
    End If
    For Each lobjstcLMI In pobjstcItem.ColecaoLMI
        If M46V999.gfPreenchido(lobjstcLMI.Nom_Logr) = False Or M46V999.gfPreenchido(lobjstcLMI.Nom_Cid) = False Or _
           M46V999.gfPreenchido(lobjstcLMI.Num_CEP) = False Or M46V999.gfPreenchido(lobjstcLMI.Sig_UF) = False Then
            Call M46V999.gpMenErro("AGE0083 - Verificar endere�os de LMI digitados na linha " & lobjstcLMI.NUM_SEQ)
        End If
    Next lobjstcLMI

SAIDA:
    GoSub FechaObj
    Exit Sub
FechaObj:
    Set lobjstcLMI = Nothing
    Return

End Sub
'************************************************************
Private Sub mpRelationCobert(ByVal pCod_Cobert, ByVal pISCobertua As Double, pListCobert() As Integer, pListISCobert() As Double, pTotCobert As Integer)
    Dim lbdTbm_Cobert_Relation As ADODB.Connection    'Arquivo: P00Multi.
    Dim lMensagem As String    'Mensagem.
    Dim lrsTbm_Cobert_Relation As clsYasRecordSet     'Registro : Tbm_Cobert_Relation.
    Dim i As Integer
    Dim lI2 As Integer
    Dim lTemCobObrigatoria As Boolean
    Dim lAchouCobObrigatoria As Boolean
    Dim lValorMaxIS As Double
    Dim pulaLimite As Boolean
    Dim lValor As Integer
    Dim lVetor As Variant
    Dim iVetor As Integer
    Dim lSelect As String

    If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, lbdTbm_Cobert_Relation, lMensagem) = False Then
        Call M46V999.gpMenErro("FGE0152 - " & lMensagem)
        Call M46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If

    'tarifa junho - inicio

    '1 - Procura pela Cobertura na Tabela de Relacionamento entre coberturas - OBRIGAT�RIAS
    lSelect = " Select *  from Tbm_Cobert_Relation WITH (NOLOCK) "
    lSelect = lSelect & " Where Cod_Ramo = " & M46V999.gCod_Ramo
    lSelect = lSelect & " and Cod_Cobert = " & pCod_Cobert
    lSelect = lSelect & " and Cod_Tarif = " & M46V999.gCodProduto
    lSelect = lSelect & " and Cod_Sequencia <= " & M46V999.gCodSequencia
    lSelect = lSelect & " and Cod_Relaction = 'O'"
    lSelect = lSelect & " and (cod_modalidade = " & M46V111.gCod_plano & " or cod_modalidade is null)"
    lSelect = lSelect & " order by Cod_Relaction"

    If M46V999.gfObtRegistro(lbdTbm_Cobert_Relation, lSelect, lrsTbm_Cobert_Relation, lMensagem) = False Then
        Call M46V999.gpFecha2(lbdTbm_Cobert_Relation)
        Call M46V999.gpMenErro("FGE0153 - " & lMensagem)
        Call M46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If
    lTemCobObrigatoria = False
    lAchouCobObrigatoria = False
    '2 - Se n�o encontrar, n�o existe restri��o para essa cobertura em rela��o a outras

    lValorMaxIS = 0
    lValor = 0
    If Not lrsTbm_Cobert_Relation.EOF Then
        '3 - Se encontrar restri��o, vai verificar cada cobertura se pode ou n�o
        For i = 1 To pTotCobert
            If pListCobert(i) <> pCod_Cobert Or pTotCobert = 1 Then
                lTemCobObrigatoria = True

                lVetor = Split(lrsTbm_Cobert_Relation("Cod_Rel_Cobert"), ",")
                For iVetor = LBound(lVetor) To UBound(lVetor)
                    If lVetor(iVetor) = pListCobert(i) And lVetor(iVetor) <> 1 Then lValor = 1
                Next

                If lValor <> 0 Then
                    If M46V999.gCod_Ramo = e_CodRamo_Empresarial Then
                        If pCod_Cobert <> M46V999.gfBuscarDePara(128, 1) And pCod_Cobert <> M46V999.gfBuscarDePara(131, 1) And _
                        pCod_Cobert <> M46V999.gfBuscarDePara(134, 1) And pCod_Cobert <> M46V999.gfBuscarDePara(139, 1) And _
                        pCod_Cobert <> M46V999.gfBuscarDePara(142, 1) And pCod_Cobert <> M46V999.gfBuscarDePara(145, 1) Then
                            If lValorMaxIS < pListISCobert(i) Then
                                lValorMaxIS = pListISCobert(i)
                            End If
                        End If
                    Else
                        If lValorMaxIS < pListISCobert(i) Then
                            lValorMaxIS = pListISCobert(i)
                        End If
                    End If
                    lValor = 0
                    lAchouCobObrigatoria = True
                End If
            End If
        Next
    End If


    If M46V999.gCodProduto < 940 Then
        lValorMaxIS = 0
    End If
    If lTemCobObrigatoria = True Then
        If lAchouCobObrigatoria = False Then
            lMensagem = "FGE0154 - A cobertura " & pCod_Cobert & " deve ser contratada com pelo menos uma das seguintes coberturas(" & lrsTbm_Cobert_Relation("Cod_Rel_Cobert") & ")."
            lMensagem = Replace(lMensagem, ",", " ou ")
            lMensagem = Replace(lMensagem, "ou )", ")")
            Call M46V999.gpMenErro(lMensagem)
            lMensagem = ""
        Else
            'De acordo com a CR84181
            If pulaLimite = False And pCod_Cobert <> 1 Then
                If M46V999.gPrograma = "D046O003" Then
                    If pISCobertua > lValorMaxIS And lValorMaxIS > 0 Then
                        lMensagem = "FGE0155 - O valor da IS da cobertura " & pCod_Cobert & " n�o pode ser maior que " & Format(lValorMaxIS, "standard") & "."
                        
                        If M46V999.gDesconsiderou <> 1 And M46V999.gPrograma <> "D046O003" Then
                            If MsgBox(lMensagem & vbCrLf & "Deseja desconsiderar esta consist�ncia?", vbQuestion + vbYesNo) = vbNo Then
                                    Call M46V999.gpMenErro(lMensagem)
                            Else
                                Call M46V999.gpDesconsiderou
                            End If
                        Else
                            Call M46V999.gpMenErro(lMensagem)
                        End If
                        lMensagem = ""
                    End If
                End If
            End If
        End If
    End If


    '4 - Procura pela Cobertura na Tabela de Relacionamento entre coberturas - PROIBIDAS
    lSelect = " Select *  from Tbm_Cobert_Relation WITH (NOLOCK) "
    lSelect = lSelect & " Where Cod_Ramo = " & M46V999.gCod_Ramo
    lSelect = lSelect & " and Cod_Cobert = " & pCod_Cobert
    lSelect = lSelect & " and Cod_Tarif = " & M46V999.gCodProduto
    lSelect = lSelect & " and Cod_Sequencia <= " & M46V999.gCodSequencia
    lSelect = lSelect & " and Cod_Relaction = 'P'"
    lSelect = lSelect & " and (cod_modalidade = " & M46V111.gCod_plano & " or cod_modalidade is null)"
    lSelect = lSelect & " order by Cod_Relaction"

    If M46V999.gfObtRegistro(lbdTbm_Cobert_Relation, lSelect, lrsTbm_Cobert_Relation, lMensagem) = False Then
        Call M46V999.gpFecha2(lbdTbm_Cobert_Relation)
        Call M46V999.gpMenErro("FGE0153 - " & lMensagem)
        Call M46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If
    lTemCobObrigatoria = False
    lAchouCobObrigatoria = False
    '5 - Se n�o encontrar, n�o existe restri��o para essa cobertura em rela��o a outras

    lValorMaxIS = 0
    lValor = 0
    If Not lrsTbm_Cobert_Relation.EOF Then
        '6 - Se encontrar restri��o, vai verificar cada cobertura se pode ou n�o
        For i = 1 To pTotCobert
            If pListCobert(i) <> pCod_Cobert Then

                lVetor = Split(lrsTbm_Cobert_Relation("Cod_Rel_Cobert"), ",")
                For iVetor = LBound(lVetor) To UBound(lVetor)
                    If lVetor(iVetor) = pListCobert(i) And lVetor(iVetor) <> 1 And pISCobertua > 0 And pListISCobert(i) > 0 Then lValor = 1
                Next

                If lValor <> 0 Then
                    If M46V999.gTipoAcesso <> DIGITACAO Then
                        Call M46V999.gpMenErro("FGE0154 - A cobertura " & pCod_Cobert & " n�o pode ser contratada com a cobertura " & pListCobert(i) & ".")
                        lValor = 0
                    End If
                End If
            End If
        Next
    End If



    'tarifa junho - at� aqui


    '4 - Procura pela Cobertura na Tabela de Relacionamento entre coberturas
    ' das coberturas que utilizam a mesma taxa

    lSelect = " Select *  from Tbm_Cob_Dep_Taxa WITH (NOLOCK) "
    lSelect = lSelect & " Where Cod_Ramo = " & M46V999.gCod_Ramo
    lSelect = lSelect & " and Cod_Modalidade = " & M46V111.gCod_plano
    lSelect = lSelect & " and Cod_Dep_Principal = " & pCod_Cobert
    lSelect = lSelect & " and Cod_Tarif = " & M46V999.gCodProduto
    lSelect = lSelect & " and Cod_Sequencia <= " & M46V999.gCodSequencia

    If M46V999.gfObtRegistro(lbdTbm_Cobert_Relation, lSelect, lrsTbm_Cobert_Relation, lMensagem) = False Then
        Call M46V999.gpFecha2(lbdTbm_Cobert_Relation)
        Call M46V999.gpMenErro("FGE0153 - " & lMensagem)
        Call M46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If

    '5 - Se n�o encontrar, essa cobertura n�o tem dependencia de taxa
    If lrsTbm_Cobert_Relation.EOF Then
        GoTo FimRotina
    End If

    Dim Achou As Boolean
    Achou = False
    '6 - Se encontrar a depend�ncia, vai verificar a cobertura principal est� contratada
    For i = 1 To pTotCobert
        If pListCobert(i) = lrsTbm_Cobert_Relation("Cod_Principal") Then
            Achou = True
            Exit For
        End If
    Next
    If Not Achou Then
        Call M46V999.gpMenErro("FGE0155 - A cobertura " & pCod_Cobert & " n�o pode ser contratada sem a cobertura " & lrsTbm_Cobert_Relation("Cod_Principal") & ".")
    End If


FimRotina:


    Call M46V999.gpFecha2(lbdTbm_Cobert_Relation)

End Sub

Private Sub mpLimiteLMG(pobjstcPed As stcA46V702B, pobjstcPedLoc As stcA46V708B)
    Dim lobjstcCobert As stcA46V704B
    Dim lMensagem As String
    Dim lretorno As String


    For Each lobjstcCobert In pobjstcPedLoc.ColecaoPedCobert
        '        Debug.Print lobjstcCobert.Cod_Cobert
        If Not M46V103.gfVerParametro(9, pobjstcPed, pobjstcPedLoc, lobjstcCobert, lretorno, lMensagem) Then
            Call M46V999.gpMenErro(lMensagem)
            GoTo SAIDA
        End If
        If M46V999.gfPreenchido(lretorno) Then
            If lobjstcCobert.Val_IS > CSng(lretorno) Then
                Call M46V999.gpMenErro("FGE0155 - A IS da cobertura " & lobjstcCobert.Cod_Cobert & " n�o pode ser maior que " & Format(CSng(lretorno), "standard"))
            End If
        End If
    Next lobjstcCobert

SAIDA:
    GoSub FechaObj
    Exit Sub
FechaObj:
    Set lobjstcCobert = Nothing
    Return
End Sub

Private Sub mpLimiteComercial(ByRef pDescComl As Double, ByRef pComissao As Double, pCodCorretor As Long, pCod_Clasf As String, pCod_Compl_Clasf As String, ByRef pPerDescLMI As Double)
    Dim lbdTbm_Limite As ADODB.Connection    'Arquivo: P00Multi.
    Dim lMensagem As String    'Mensagem.
    Dim lrsTbm_Limite As clsYasRecordSet    'Registro : Tbm_Limite.
    Dim i As Integer
    Dim lSelect As String

    If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, lbdTbm_Limite, lMensagem) = False Then
        Call M46V999.gpMenErro("FGE0152 - " & lMensagem)
        Call M46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If

    pDescComl = 0
    pComissao = 0

    '1 - Procura pelo corretor na Tabela de Limite dos Corretores
    'Desconto Comercial para esse corretor - Exce��o
    lSelect = " Select *  from Tbm_Limite WITH (NOLOCK) "
    lSelect = lSelect & " Where Cod_Ramo = " & M46V999.gCod_Ramo
    lSelect = lSelect & " and Cod_Modalidade = " & M46V111.gCod_plano
    lSelect = lSelect & " and Cod_Tipo_Reg = 'E'"
    lSelect = lSelect & " and Cod_Tip_Limite = '1' "
    lSelect = lSelect & " and Cod_Corr = " & pCodCorretor
    lSelect = lSelect & " and Cod_Tarif = " & M46V999.gCodProduto
    lSelect = lSelect & " and Cod_Sequencia <= " & M46V999.gCodSequencia

    If M46V999.gfObtRegistro(lbdTbm_Limite, lSelect, lrsTbm_Limite, lMensagem) = False Then
        Call M46V999.gpFecha2(lbdTbm_Limite)
        Call M46V999.gpMenErro("FGE0153 - " & lMensagem)
        Call M46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If

    '2 - Se n�o encontrar, vai procurar default
    If Not lrsTbm_Limite.EOF Then
        pDescComl = M46V999.CDblx(Replace(Trim(lrsTbm_Limite("Val_Limite")), ".", ",")) 'M46V999.CDblx(lrsTbm_Limite("Val_Limite"))
        If val(pCod_Clasf) = M46V999.gfBuscarDePara(22, 2) And val(pCod_Compl_Clasf) = M46V999.gfBuscarDePara(22, 2, 13) Then
            Select Case pCodCorretor
            Case 49301, 8001, 8002, 8003, 8004, 8005, 8006, 8008, 8011, 8014, 8028, _
                 3401, 345601, 486001, 489501, 646401, 5801, 648701, 648702, _
                 648705, 648706, 648708, 358301, 219101, 15701, 493701, _
                 6901, 686601, 403801, 8101
                pDescComl = 40
            End Select
        End If
    Else

        'Desconto Comercial para qualquer corretor - Default
        lSelect = " Select *  from Tbm_Limite WITH (NOLOCK) "
        lSelect = lSelect & " Where Cod_Ramo = " & M46V999.gCod_Ramo
        lSelect = lSelect & " and Cod_Modalidade = " & M46V111.gCod_plano
        lSelect = lSelect & " and Cod_Tipo_Reg = 'D'"
        lSelect = lSelect & " and Cod_Tip_Limite = '1' "
        lSelect = lSelect & " and Cod_Tarif = " & M46V999.gCodProduto
        lSelect = lSelect & " and Cod_Sequencia <= " & M46V999.gCodSequencia

        If M46V999.gfObtRegistro(lbdTbm_Limite, lSelect, lrsTbm_Limite, lMensagem) = False Then
            Call M46V999.gpFecha2(lbdTbm_Limite)
            Call M46V999.gpMenErro("FGE0153 - " & lMensagem)
            Call M46V999.gpGraLog(2, lMensagem)
            Exit Sub
        End If
        If Not lrsTbm_Limite.EOF Then
            pDescComl = M46V999.CDblx(Replace(Trim(lrsTbm_Limite("Val_Limite")), ".", ",")) 'M46V999.CDblx(lrsTbm_Limite("Val_Limite"))
        Else
            pDescComl = 0
        End If
    End If

    'Segmentados (Permitir Limite de comiss�o especifico (E) por sistemas (Web / Kit))
    If M46V999.gOrigemChamada = e_Origem_Kit And M46V111.gCod_plano > 5000 Then
        pCodCorretor = 0
    End If

    '3 - Procura pelo corretor na Tabela de Limite dos Corretores
    'Limite de Comiss�o para esse corretor - Exce��o
    lSelect = " Select *  from Tbm_Limite WITH (NOLOCK) "
    lSelect = lSelect & " Where Cod_Ramo = " & M46V999.gCod_Ramo
    lSelect = lSelect & " and Cod_Modalidade = " & M46V111.gCod_plano
    lSelect = lSelect & " and Cod_Tipo_Reg = 'E'"
    lSelect = lSelect & " and Cod_Tip_Limite = '2' "
    lSelect = lSelect & " and Cod_Corr = " & pCodCorretor
    lSelect = lSelect & " and Cod_Tarif = " & M46V999.gCodProduto
    lSelect = lSelect & " and Cod_Sequencia <= " & M46V999.gCodSequencia

    If M46V999.gfObtRegistro(lbdTbm_Limite, lSelect, lrsTbm_Limite, lMensagem) = False Then
        Call M46V999.gpFecha2(lbdTbm_Limite)
        Call M46V999.gpMenErro("FGE0153 - " & lMensagem)
        Call M46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If

    '4 - Se n�o encontrar, vai procurar default
    If Not lrsTbm_Limite.EOF Then
        pComissao = M46V999.CDblx(Replace(Trim(lrsTbm_Limite("Val_Limite")), ".", ",")) 'M46V999.CDblx(lrsTbm_Limite("Val_Limite"))
    Else

        'Desconto Comercial para qualquer corretor - Default
        lSelect = " Select *  from Tbm_Limite WITH (NOLOCK) "
        lSelect = lSelect & " Where Cod_Ramo = " & M46V999.gCod_Ramo
        lSelect = lSelect & " and Cod_Modalidade = " & M46V111.gCod_plano
        lSelect = lSelect & " and Cod_Tipo_Reg = 'D'"
        lSelect = lSelect & " and Cod_Tip_Limite = '2' "
        lSelect = lSelect & " and Cod_Tarif = " & M46V999.gCodProduto
        lSelect = lSelect & " and Cod_Sequencia <= " & M46V999.gCodSequencia

        If M46V999.gfObtRegistro(lbdTbm_Limite, lSelect, lrsTbm_Limite, lMensagem) = False Then
            Call M46V999.gpFecha2(lbdTbm_Limite)
            Call M46V999.gpMenErro("FGE0153 - " & lMensagem)
            Call M46V999.gpGraLog(2, lMensagem)
            Exit Sub
        End If
        If Not lrsTbm_Limite.EOF Then
            pComissao = M46V999.CDblx(Replace(Trim(lrsTbm_Limite("Val_Limite")), ".", ",")) 'M46V999.CDblx(lrsTbm_Limite("Val_Limite"))
        Else
            pComissao = 0
        End If
    End If

    '5-Verifica se � planinho

    '1000744 - Item 4 - Plano Mercedes
    If (M46V111.gCod_plano >= 2021 And M46V111.gCod_plano <= 2031) And (M46V999.gCodProduto < 992) Then
        If Trim(pCod_Clasf) = M46V999.gfBuscarDePara("001", 2) And Trim(pCod_Compl_Clasf) = M46V999.gfBuscarDePara("001", 2, "11") Then
            pDescComl = 33
        End If
        If Trim(pCod_Clasf) = M46V999.gfBuscarDePara("001", 2) And Trim(pCod_Compl_Clasf) = M46V999.gfBuscarDePara("001", 2, "21") Then
            pDescComl = 20
        End If
    End If

    'Ficha 901673 e 901772
    If (M46V111.gCod_plano >= 2051 And M46V111.gCod_plano <= 2100) Or (M46V111.gCod_plano >= 2111 And M46V111.gCod_plano <= 2130) Then
        If Trim(pCod_Clasf) = M46V999.gfBuscarDePara("001", 2) And Trim(pCod_Compl_Clasf) = M46V999.gfBuscarDePara("001", 2, "11") Then
            pDescComl = 20
        End If
        '1200903 - Tarifa - 1046 e 1047 - Item 13
        If (M46V111.gCod_plano >= 2051 And M46V111.gCod_plano <= 2060 And M46V999.gCodProduto < 1046) Then
            If Trim(pCod_Clasf) = M46V999.gfBuscarDePara("001", 2) And Trim(pCod_Compl_Clasf) = M46V999.gfBuscarDePara("001", 2, "21") Then
                pDescComl = 5
            End If
        End If
    End If

    '1200903 - Tarifa - 1046 e 1047 - Item 13
    If (M46V111.gCod_plano >= 2051 And M46V111.gCod_plano <= 2060) And (M46V999.gCodProduto < 1046) Then
        pDescComl = 0
    End If

    If M46V999.gNovaPropostaLMI Then
        '6 - Procura pelo corretor na Tabela de Limite dos Corretores
        'Limite de Desconto de LMI para esse corretor - Exce��o
        lSelect = " Select *  from Tbm_Limite WITH (NOLOCK) "
        lSelect = lSelect & " Where Cod_Ramo = " & M46V999.gCod_Ramo
        lSelect = lSelect & " and Cod_Modalidade = " & M46V111.gCod_plano
        lSelect = lSelect & " and Cod_Tipo_Reg = 'E'"
        lSelect = lSelect & " and Cod_Tip_Limite = '3' "
        lSelect = lSelect & " and Cod_Corr = " & pCodCorretor
        lSelect = lSelect & " and Cod_Tarif = " & M46V999.gCodProduto
        lSelect = lSelect & " and Cod_Sequencia <= " & M46V999.gCodSequencia

        If M46V999.gfObtRegistro(lbdTbm_Limite, lSelect, lrsTbm_Limite, lMensagem) = False Then
            Call M46V999.gpFecha2(lbdTbm_Limite)
            Call M46V999.gpMenErro("FGE0156 - " & lMensagem)
            Call M46V999.gpGraLog(2, lMensagem)
            Exit Sub
        End If

        '7 - Se n�o encontrar, vai procurar default
        If Not lrsTbm_Limite.EOF Then
            pPerDescLMI = M46V999.CDblx(Replace(Trim(lrsTbm_Limite("Val_Limite")), ".", ",")) ' M46V999.CDblx(lrsTbm_Limite("Val_Limite"))

        Else

            'Desconto Comercial para qualquer corretor - Default
            lSelect = " Select *  from Tbm_Limite WITH (NOLOCK) "
            lSelect = lSelect & " Where Cod_Ramo = " & M46V999.gCod_Ramo
            lSelect = lSelect & " and Cod_Modalidade = " & M46V111.gCod_plano
            lSelect = lSelect & " and Cod_Tipo_Reg = 'D'"
            lSelect = lSelect & " and Cod_Tip_Limite = '3' "
            lSelect = lSelect & " and Cod_Tarif = " & M46V999.gCodProduto
            lSelect = lSelect & " and Cod_Sequencia <= " & M46V999.gCodSequencia

            If M46V999.gfObtRegistro(lbdTbm_Limite, lSelect, lrsTbm_Limite, lMensagem) = False Then
                Call M46V999.gpFecha2(lbdTbm_Limite)
                Call M46V999.gpMenErro("FGE0157 - " & lMensagem)
                Call M46V999.gpGraLog(2, lMensagem)
                Exit Sub
            End If
            If Not lrsTbm_Limite.EOF Then
                pPerDescLMI = M46V999.CDblx(Replace(Trim(lrsTbm_Limite("Val_Limite")), ".", ",")) 'M46V999.CDblx(lrsTbm_Limite("Val_Limite"))
            Else
                pPerDescLMI = 0
            End If
        End If

    End If

FimRotina:

    Call M46V999.gpFecha2(lbdTbm_Limite)
    Set lrsTbm_Limite = Nothing
End Sub

Private Sub mpConsisteBens(TipChamada As String, pobjstcItem As stcA46V708B)
    Dim lobjstcPedBem As stcA46V715B      'Objeto da classe de estrutura stcA46V715B (TAB_PED_LOC_BEM).

    For Each lobjstcPedBem In pobjstcItem.ColecaoPedLocBem
        'If M46V999.gTipUsuario = e_Tipo_Usu_Corretor Then
        If TipChamada = 1 Then    ' Consist�ncia
            If lobjstcPedBem.gClasseBem = "3" Then
                Call M46V999.gpMenErro("FGB0001 - " & lobjstcPedBem.gObservacao)
            End If
        End If
        If TipChamada = 2 Then    ' Impress�o
            If lobjstcPedBem.gClasseBem = "4" Then
                Call M46V999.gpMenErro("FGB0001 - " & lobjstcPedBem.gObservacao)
            End If
        End If
        'Else
        '    If lobjstcPedBem.gClasseBem = 3 Or lobjstcPedBem.gClasseBem = 4 Then
        ''' GravaPropMensagem
        ''Suspender
        '    End If
        'End If
    Next
    Set lobjstcPedBem = Nothing
End Sub

'rodrigoS
'Precisa retirar da DLL
'Private Sub mpConsisteCapitalGlobal()
'    If M46V999.gfDisponivel("frmT46V181A") Then
'        With frmT46V181A
'            If frmT46V101A.optVida(0).Value = True Then
'                If Not M46V999.gfPreenchido(.cboCobertura.Text) Then
'                    Call M46V999.gpMenErro("FGE0112 - Cobertura obrigat�ria se selecionar plano de vida.")
'                End If
'                If Not M46V999.gfPreenchido(.txtQtdeVidas.Text) Then
'                    Call M46V999.gpMenErro("FGE0112 - Qtde. vidas obrigat�ria se selecionar plano de vida.")
'                End If
'                If Not M46V999.gfPreenchido(.txtLMGPorVida.Text) Then
'                    Call M46V999.gpMenErro("FGE0112 - LMG por vida obrigat�rio se selecionar plano de vida.")
'                End If
'            End If
'        End With
'    End If
'End Sub

Private Sub mpConsisteLucrosCessantes(pobjstcItem As stcA46V708B, pobjstcPed As stcA46V702B)
    Dim lStringCoberturasContratadas As String
    Dim lTotLMGCoberturasContratadas As Double
    Dim lbdTbm_Coef_LMG_VR As ADODB.Connection
    Dim lrsTbm_Coef_LMG_VR As clsYasRecordSet
    Dim lMensagem As String
    Dim lSelect As String
    Dim lobjstcCobert As stcA46V704B
    Dim lPerc_LMG_VR As Double

    '----------------------------------------
    'Pesquisa grid de coberturas contratadas.
    '----------------------------------------
    lTotLMGCoberturasContratadas = 0
    lStringCoberturasContratadas = Empty
    For Each lobjstcCobert In pobjstcItem.ColecaoPedCobert
        
        If lobjstcCobert.gCod_Grupo_Cobert = "3" Then
            lTotLMGCoberturasContratadas = lTotLMGCoberturasContratadas + lobjstcCobert.Val_IS
            If lStringCoberturasContratadas = Empty Then
                lStringCoberturasContratadas = val(lobjstcCobert.Cod_Cobert)
            Else
                lStringCoberturasContratadas = lStringCoberturasContratadas & ", " & val(lobjstcCobert.Cod_Cobert)
            End If
        End If
    Next

    If pobjstcItem.Val_Vr_Lc < lTotLMGCoberturasContratadas Then
        Call M46V999.gpMenErro("FGE0112 - Valor do risco Lucros Cessantes B�sica ou Despesas Fixas B�sica n�o pode ser < total dos LMI's das coberturas " & lStringCoberturasContratadas & ".")
        GoTo Fim
    End If

    If lTotLMGCoberturasContratadas < pobjstcItem.Val_Vr_Lc * 0.05 Then
        Call M46V999.gpMenErro("FGE0112 - O valor da somat�ria das coberturas " & lStringCoberturasContratadas & " (LMI) n�o pode ser inferior a 5% do VR de Lucros Cessantes.")
        GoTo Fim
    End If

    If CDbl(pobjstcItem.Val_Vr_Lc) > 0 Then
        lPerc_LMG_VR = lTotLMGCoberturasContratadas / CDbl(pobjstcItem.Val_Vr_Lc) * 100

        'inicio
        If M46V999.gfAbrBasDados("P00MULTI", M46V999.gTipP00MULTI, M46V999.gEndP00MULTI, lbdTbm_Coef_LMG_VR, lMensagem) = False Then
            Call M46V999.gpGraLog(2, lMensagem)
            GoTo Fim
        End If

        lSelect = " SELECT TOP 1 *"
        lSelect = lSelect & " FROM Tbm_Coef_LMG_VR"
        lSelect = lSelect & " WHERE Cod_Ramo = " & e_CodRamo_Empresarial
        lSelect = lSelect & " AND Cod_Grupo_Cobert = 3"
        lSelect = lSelect & " AND Cod_Tarif = " & M46V999.gCodProduto
        lSelect = lSelect & " AND Perc_LMI_VR <= " & Replace(lPerc_LMG_VR, ",", ".")
        lSelect = lSelect & " ORDER BY Perc_LMI_VR DESC"

        If M46V999.gfObtRegistro(lbdTbm_Coef_LMG_VR, lSelect, lrsTbm_Coef_LMG_VR, lMensagem) = False Then
            Call M46V999.gpFecha2(lbdTbm_Coef_LMG_VR)
            Call M46V999.gpGraLog(2, lMensagem)
            GoTo Fim
        End If

        If lrsTbm_Coef_LMG_VR.EOF Then
            Call M46V999.gpFecha2(lbdTbm_Coef_LMG_VR)
            Call M46V999.gpGraLog(2, "FGE0112 - Tbm_Coef_LMG_VR.Perc_LMI_VR maior ou igual a " & lPerc_LMG_VR & " n�o encontrado.")
            GoTo Fim
        End If
    End If

Fim:
    Call M46V999.gpFecha2(lbdTbm_Coef_LMG_VR)
    Set lobjstcCobert = Nothing
End Sub

Private Sub mpConsisteCondominioDanosEletricos(pobjstcItem As stcA46V708B)
    Dim Item As stcA46V704B
    
    For Each Item In pobjstcItem.ColecaoPedCobert
        If Item.Cod_Cobert = 1103 Then
            If pobjstcItem.Possui_Elevador = 0 Then
                Call M46V999.gpMenErro("Consist�ncia do comum - � necess�rio informar se possui elevador.")
            End If
        End If
    Next

    Set Item = Nothing
End Sub

Private Sub mpConsisteRelacaoBicicletaNaoInformada(pobjstcItem As stcA46V708B)
    Dim Item As stcA46V704B
    
    If pobjstcItem.ColecaoPedLocBicicleta.Count = 0 Then
        For Each Item In pobjstcItem.ColecaoPedCobert
            If Item.Cod_Cobert = 1215 Then
                Call M46V999.gpMenErro("Consist�ncia do comum - � necess�rio informar a rela��o de Bicicletas.")
            End If
        Next
    End If

    Set Item = Nothing
End Sub

Private Sub mpConsisteCoberturaBicicletaNaoInformada(pobjstcItem As stcA46V708B)
    Dim Item As stcA46V704B
    Dim blnAchou As Boolean
    
    blnAchou = False
    If pobjstcItem.ColecaoPedLocBicicleta.Count > 0 Then
        For Each Item In pobjstcItem.ColecaoPedCobert
            If Item.Cod_Cobert = 1215 Then
                blnAchou = True
                Exit For
            End If
        Next
        If Not blnAchou Then
            Call M46V999.gpMenErro("Consist�ncia do comum - � necess�rio adicionar a cobertura de Bicicleta.")
        End If
    End If
    Set Item = Nothing
End Sub

Private Sub mpConsisteValorBicicleta(pobjstcItem As stcA46V708B)
    Dim Item As stcA46V744B
    
    For Each Item In pobjstcItem.ColecaoPedLocBicicleta
        If Item.Valor = 0 Then
            Call M46V999.gpMenErro("Consist�ncia do comum - � necess�rio informar o valor do item " & Item.Cod_Seq & " da rela��o de bicicletas.")
        End If
    Next
    Set Item = Nothing
End Sub

Private Sub mpConsisteDescricaoBicicleta(pobjstcItem As stcA46V708B)
    Dim Item As stcA46V744B
    
    For Each Item In pobjstcItem.ColecaoPedLocBicicleta
        If Not M46V999.gfPreenchido(Item.Descricao) Then
            Call M46V999.gpMenErro("Consist�ncia do comum - � necess�rio informar uma descri��o para o item " & Item.Cod_Seq & " da rela��o de bicicletas.")
        End If
    Next
    Set Item = Nothing
End Sub
Private Sub mpConsisteRelacaoAllRisksNaoInformada(pobjstcItem As stcA46V708B)
    Dim Item As stcA46V704B
    
    If pobjstcItem.ColecaoAllRisks.Count = 0 Then
        For Each Item In pobjstcItem.ColecaoPedCobert
            If Item.Cod_Cobert = 1172 Then
                Call M46V999.gpMenErro("Consist�ncia do comum - � necess�rio informar a rela��o de All Risks.")
            End If
        Next
    End If

    Set Item = Nothing
End Sub

Private Sub mpConsisteCoberturaAllRisksNaoInformada(pobjstcItem As stcA46V708B)
    Dim Item As stcA46V704B
    Dim blnAchou As Boolean
    
    blnAchou = False
    If pobjstcItem.ColecaoAllRisks.Count > 0 Then
        For Each Item In pobjstcItem.ColecaoPedCobert
            If Item.Cod_Cobert = 1172 Then
                blnAchou = True
                Exit For
            End If
        Next
        If Not blnAchou Then
            Call M46V999.gpMenErro("Consist�ncia do comum - � necess�rio adicionar a cobertura de All Risks.")
        End If
    End If
    Set Item = Nothing
End Sub

Private Sub mpConsisteValorAllRisks(pobjstcItem As stcA46V708B)
    Dim Item As stcA46V747B
    
    For Each Item In pobjstcItem.ColecaoAllRisks
        If Item.Valor = 0 Then
            Call M46V999.gpMenErro("Consist�ncia do comum - � necess�rio informar o valor do item " & Item.Cod_Seq & " da rela��o de allRisks.")
        End If
    Next
    Set Item = Nothing
End Sub

Private Sub mpConsisteDescricaoAllRisks(pobjstcItem As stcA46V708B)
    Dim Item As stcA46V747B
    
    For Each Item In pobjstcItem.ColecaoAllRisks
        If Not M46V999.gfPreenchido(Item.Descricao) Then
            Call M46V999.gpMenErro("Consist�ncia do comum - � necess�rio informar uma descri��o para o item " & Item.Cod_Seq & " da rela��o de allRisks.")
        End If
    Next
    Set Item = Nothing
End Sub

Private Sub mpConsisteRelacaoObrasArteNaoInformada(pobjstcItem As stcA46V708B)
    Dim Item As stcA46V704B
    
    If pobjstcItem.ColecaoObrasArte.Count = 0 Then
        For Each Item In pobjstcItem.ColecaoPedCobert
            If Item.Cod_Cobert = 3022 Then
                Call M46V999.gpMenErro("Consist�ncia do comum - � necess�rio informar a rela��o de Obras de arte.")
            End If
        Next
    End If

    Set Item = Nothing
End Sub

Private Sub mpConsisteCoberturaObrasArteNaoInformada(pobjstcItem As stcA46V708B)
    Dim Item As stcA46V704B
    Dim blnAchou As Boolean
    
    blnAchou = False
    If pobjstcItem.ColecaoObrasArte.Count > 0 Then
        For Each Item In pobjstcItem.ColecaoPedCobert
            If Item.Cod_Cobert = 3022 Then
                blnAchou = True
                Exit For
            End If
        Next
        If Not blnAchou Then
            Call M46V999.gpMenErro("Consist�ncia do comum - � necess�rio adicionar a cobertura de Obras de arte.")
        End If
    End If
    Set Item = Nothing
End Sub

Private Sub mpConsisteValorObrasArte(pobjstcItem As stcA46V708B)
    Dim Item As stcA46V748B
    
    For Each Item In pobjstcItem.ColecaoObrasArte
        If Item.Valor = 0 Then
            Call M46V999.gpMenErro("Consist�ncia do comum - � necess�rio informar o valor do item " & Item.Cod_Seq & " da rela��o de Obras de arte.")
        End If
    Next
    Set Item = Nothing
End Sub

Private Sub mpConsisteDescricaoObrasArte(pobjstcItem As stcA46V708B)
    Dim Item As stcA46V748B
    
    For Each Item In pobjstcItem.ColecaoObrasArte
        If Not M46V999.gfPreenchido(Item.Descricao) Then
            Call M46V999.gpMenErro("Consist�ncia do comum - � necess�rio informar uma descri��o para o item " & Item.Cod_Seq & " da rela��o de Obras de arte.")
        End If
    Next
    Set Item = Nothing
End Sub
Private Sub mpConsistePessoa(pobjstcItem As stcA46V708B)
    Dim lobjstcPedPessoa As stcA46V735B
    Dim clsA99V001 As clsA99V001

    Dim lDig As String
    Dim lAuxiliar As String
    Dim lMensagem As String

    Set clsA99V001 = New clsA99V001
    Set clsA99V001.M46V999 = M46V999

    For Each lobjstcPedPessoa In pobjstcItem.ColecaoBeneficiario
        With lobjstcPedPessoa
            If .Tip_Pessoa <> 1 And .Tip_Pessoa <> 2 And .Tip_Pessoa <> 3 Then
                Call M46V999.gpMenErro("CSP0001 - Benefici�rio deve ser Pessoa F�sica, Jur�dica ou Institui��o Financeira.")
            End If
            If Trim(.Nom_Pessoa) = "" Then
                Call M46V999.gpMenErro("CSP0002 - Nome Benefici�rio deve ser preenchido.")
            End If
            If .Tip_Cobert > 2 Or .Tip_Cobert < 0 Then
                Call M46V999.gpMenErro("CSP0003 - Benefici�rio - Bens Compreendidos deve ser preenchido.")
            End If
            If .Per_Participacao = 0 Then
                Call M46V999.gpMenErro("CSP0004 - Benefici�rio - Percentual de Participa��o deve ser preenchido.")
            End If
            If .Tip_Pessoa = 1 Or .Tip_Pessoa = 3 Then
                Call M46V999.gpForCNPJ(1, .Num_CPF_CNPJ, lAuxiliar)
                If Not clsA99V001.DIG_CNPJ(lAuxiliar, lDig, lMensagem) Then
                    Call M46V999.gpMenErro("CSP0005 - Benefici�rio - N�mero do CNPJ inv�lido.")
                End If
            ElseIf .Tip_Pessoa = 2 Then
                Call M46V999.gpForCPF(1, .Num_CPF_CNPJ, lAuxiliar)
                If Not clsA99V001.DIG_CPF(lAuxiliar, lDig, lMensagem) Then
                    Call M46V999.gpMenErro("CSP0006 - Benefici�rio - N�mero do CPF inv�lido.")
                End If
            End If
        End With
    Next

    With pobjstcItem.objstcPedInquilino
        If M46V999.gfPreenchido(.Num_CPF_CNPJ) And Trim(.Num_CPF_CNPJ) <> "0" Then
            If .Tip_Pessoa <> 1 And .Tip_Pessoa <> 2 Then
                Call M46V999.gpMenErro("CSP0007 - Inquilino deve ser Pessoa F�sica ou Jur�dica")
            End If
            If Trim(.Nom_Pessoa) = "" Then
                Call M46V999.gpMenErro("CSP0008 - Nome Inquilino deve ser preenchido.")
            End If
            If .Tip_Pessoa = 1 Then
                Call M46V999.gpForCNPJ(1, .Num_CPF_CNPJ, lAuxiliar)
                If Not clsA99V001.DIG_CNPJ(lAuxiliar, lDig, lMensagem) Then
                    Call M46V999.gpMenErro("CSP0009 - Inquilino - N�mero do CNPJ inv�lido.")
                End If
            ElseIf .Tip_Pessoa = 2 Then
                Call M46V999.gpForCPF(1, .Num_CPF_CNPJ, lAuxiliar)
                If Not clsA99V001.DIG_CPF(lAuxiliar, lDig, lMensagem) Then
                    Call M46V999.gpMenErro("CSP0010 - Inquilino - N�mero do CPF inv�lido.")
                End If
            End If
        End If
    End With


    With pobjstcItem.objstcPedProprietario
        If M46V999.gfPreenchido(.Num_CPF_CNPJ) And Trim(.Num_CPF_CNPJ) <> "0" Then
            If .Tip_Pessoa <> 1 And .Tip_Pessoa <> 2 Then
                Call M46V999.gpMenErro("CSP0011 - Propriet�rio deve ser Pessoa F�sica ou Jur�dica")
            End If
            If Trim(.Nom_Pessoa) = "" Then
                Call M46V999.gpMenErro("CSP0012 - Nome Propriet�rio deve ser preenchido.")
            End If
            If .Tip_Pessoa = 1 Then
                Call M46V999.gpForCNPJ(1, .Num_CPF_CNPJ, lAuxiliar)
                If Not clsA99V001.DIG_CNPJ(lAuxiliar, lDig, lMensagem) Then
                    Call M46V999.gpMenErro("CSP0013 - Propriet�rio - N�mero do CNPJ inv�lido.")
                End If
            ElseIf .Tip_Pessoa = 2 Then
                Call M46V999.gpForCPF(1, .Num_CPF_CNPJ, lAuxiliar)
                If Not clsA99V001.DIG_CPF(lAuxiliar, lDig, lMensagem) Then
                    Call M46V999.gpMenErro("CSP0014 - Propriet�rio - N�mero do CPF inv�lido.")
                End If
            End If
        End If
    End With

    Set lobjstcPedPessoa = Nothing

End Sub

Private Sub mpConsisteProlabore(ColecaoProlabore As Collection, ByVal pComissaoCorretor As Double, ByVal pCod_Tip_Doc As Integer)
    Dim lobjstcProlabore As stcA46V733B
    Dim lPercLimite As Double
    Dim lTotalPerc As Double
    Dim lMensagem As String
    Dim StringProlabore As String

    On Local Error GoTo ControleErro

    lTotalPerc = 0
    StringProlabore = ""

    For Each lobjstcProlabore In ColecaoProlabore
        With lobjstcProlabore
            If M46V999.gfObtLimiteMaximoProlaboreCorretor2(.Cod_Canal, lPercLimite, lMensagem) Then
                lTotalPerc = lTotalPerc + .Perc_Canal
                If .Perc_Canal > lPercLimite Then
                    Call M46V999.gpMenErro("FGE5003 - Limite Prolabore Inv�lido (Acima do Limite Permitido).")
                End If
            Else
                Call M46V999.gpMenErro("CSP0005 - " & lMensagem)
            End If
            If StringProlabore <> "" And e_TipDoc_Cotacao <> pCod_Tip_Doc Then
                If InStr(StringProlabore, "@" & .Cod_Canal & "@") > 0 Then
                    Call M46V999.gpMenErro("CSP0006 - Mesmo prolaborista n�o pode ser selecionado mais de uma vez.")
                End If
            End If
            StringProlabore = StringProlabore & ",@" & .Cod_Canal & "@"

        End With
    Next lobjstcProlabore

    Set lobjstcProlabore = Nothing

    If lTotalPerc + pComissaoCorretor > 99.9 Then
        Call M46V999.gpMenErro("CSP0006 - Soma das comiss�es dos Corretores e Prolaboristas n�o pode ser maior do que 99.9 %")
    End If

    Exit Sub
ControleErro:

    lMensagem = "mpConsisteProlabore Erro - " & Err & " " & Err.Description
    Call M46V999.gpGraLog(0, lMensagem)

End Sub

Private Sub mpConsisteLimiteAgravo(pobjstcCtrlEmis As stcA46V712B)
    Dim lPercLimite As Double
    Dim lMensagem As String

    On Local Error GoTo ControleErro

    If pobjstcCtrlEmis.Per_Desc_Agrav = 0 Then Exit Sub

    If M46V111.gCod_plano > 1000 Then
        Call M46V999.gpMenErro("FGR0002 - " & "Agravo n�o pode ser aplicado aos planos simplificados.")
        Exit Sub
    End If

    If M46V999.gfObtLimiteMaximoAgravo(pobjstcCtrlEmis.Cod_Corr, lPercLimite, lMensagem) Then
        If pobjstcCtrlEmis.Per_Desc_Agrav > lPercLimite Then
            Call M46V999.gpMenErro("FGR0003 - Limite de Agravo Inv�lido (Acima do Limite Permitido).")
            pobjstcCtrlEmis.Per_Desc_Agrav = 0    'Se n�o zerar n�o ser� possivel recalcular
        End If
    Else
        Call M46V999.gpMenErro("FGR0004 - " & lMensagem)
    End If

    Exit Sub
ControleErro:

    lMensagem = "mpConsisteLimiteAgravo Erro - " & Err & " " & Err.Description
    Call M46V999.gpGraLog(0, lMensagem)

End Sub



Private Sub mpConsisteFatoresProtecao(pobjstcItem As stcA46V708B)
    Dim lobjstcFator As stcA46V720B
    Dim lAux As String
    Dim TemErro As Boolean
    lAux = ""

    If pobjstcItem.Cod_Ramo <> e_CodRamo_Empresarial Then
        Exit Sub
    End If
    For Each lobjstcFator In pobjstcItem.ColecaoPedFator
        lAux = lAux & ";" & lobjstcFator.Cod_Fator & ";"
    Next

    TemErro = False

    For Each lobjstcFator In pobjstcItem.ColecaoPedFator
        If lobjstcFator.Cod_Fator = 1 Then
            If InStr(lAux, ";9;") > 0 Then
                Call M46V999.gpMenErro("FFP0001 - Prote��o: 'Extintor' n�o pode ser contratado com 'Risco em Shopping Center'")
                TemErro = True
            End If
        End If
        If lobjstcFator.Cod_Fator = 2 Then
            If InStr(lAux, ";9;") > 0 Then
                Call M46V999.gpMenErro("FFP0002 - Prote��o: 'Hidrantes' n�o pode ser contratado com 'Risco em Shopping Center'")
                TemErro = True
            End If
        End If
        If lobjstcFator.Cod_Fator = 3 Then
            If InStr(lAux, ";4;") > 0 Or InStr(lAux, ";5;") > 0 Or InStr(lAux, ";6;") > 0 Or InStr(lAux, ";9;") > 0 Then
                Call M46V999.gpMenErro("FFP0003 - Prote��o: 'Outros exceto Sprinklers' n�o pode ser contratado com 'Sprinklers, conforme TSIB' e/ou 'Detec��o e Alarme' e/ou 'Detec��o e Alarme com Sprinklers' e/ou 'Risco em Shopping Center'")
                TemErro = True
            End If
        End If
        If lobjstcFator.Cod_Fator = 4 Then
            If InStr(lAux, ";3;") > 0 Or InStr(lAux, ";5;") > 0 Or InStr(lAux, ";6;") > 0 Or InStr(lAux, ";9;") > 0 Then
                Call M46V999.gpMenErro("FFP0004 - Prote��o: 'Sprinklers, conforme TSIB' n�o pode ser contratado com 'Outros exceto Sprinklers' e/ou 'Detec��o e Alarme'  e/ou 'Detec��o e Alarme com Sprinklers' e/ou 'Risco em Shopping Center'")
                TemErro = True
            End If
        End If
        If lobjstcFator.Cod_Fator = 5 Then
            If InStr(lAux, ";3;") > 0 Or InStr(lAux, ";4;") > 0 Or InStr(lAux, ";6;") > 0 Or InStr(lAux, ";9;") > 0 Then
                Call M46V999.gpMenErro("FFP0005 - Prote��o: 'Detec��o e Alarme' n�o pode ser contratado com 'Outros exceto Sprinklers' e/ou 'Sprinklers, conforme TSIB'  e/ou 'Detec��o e Alarme com Sprinklers' e/ou 'Risco em Shopping Center'")
                TemErro = True
            End If
        End If
        If lobjstcFator.Cod_Fator = 6 Then
            If InStr(lAux, ";3;") > 0 Or InStr(lAux, ";4;") > 0 Or InStr(lAux, ";5;") > 0 Or InStr(lAux, ";9;") > 0 Then
                Call M46V999.gpMenErro("FFP0006 - Prote��o: 'Detec��o e Alarme com Sprinklers' n�o pode ser contratado com 'Outros exceto Sprinklers' e/ou 'Sprinklers, conforme TSIB'  e/ou 'Detec��o e Alarme' e/ou 'Risco em Shopping Center'")
                TemErro = True
            End If
        End If
        If lobjstcFator.Cod_Fator = 7 Then
            If InStr(lAux, ";9;") > 0 Then
                Call M46V999.gpMenErro("FFP0007 - Prote��o: 'Alarme' n�o pode ser contratado com 'Risco em Shopping Center'")
                TemErro = True
            End If
        End If
        If lobjstcFator.Cod_Fator = 8 Then
            If InStr(lAux, ";9;") > 0 Then
                Call M46V999.gpMenErro("FFP0008 - Prote��o: 'Vigil�ncia Armada' n�o pode ser contratado com 'Risco em Shopping Center'")
                TemErro = True
            End If
        End If
        If lobjstcFator.Cod_Fator = 9 Then
            If InStr(lAux, ";1;") > 0 Or InStr(lAux, ";2;") > 0 Or InStr(lAux, ";3;") > 0 Or InStr(lAux, ";4;") > 0 Or InStr(lAux, ";5;") > 0 Or InStr(lAux, ";6;") > 0 Or InStr(lAux, ";7;") > 0 Or InStr(lAux, ";8;") > 0 Then
                Call M46V999.gpMenErro("FFP0009 - Prote��o: 'Risco em Shopping Center' deve ser contratado sozinho!")
                TemErro = True
            End If
        End If

        If TemErro Then
            Exit For
        End If
    Next
    Set lobjstcFator = Nothing
End Sub

Private Sub mpVerificarInspecao(pobjstcItem As stcA46V708B, pobjstcPed As stcA46V702B)
    Dim lSelect As String
    Dim lrst As clsYasRecordSet
    Dim bdGed As ADODB.Connection
    Dim lMensagem As String

    If pobjstcItem.Num_Insp = 0 Then
        Exit Sub
    End If

    If Not M46V999.gfAbrBasDados("GED", M46V999.gTipGed, M46V999.gEndGed, bdGed, lMensagem) Then
        Call M46V999.gpGraLog(1, lMensagem)
        Exit Sub
    End If

    lSelect = " Select * from inspecao.dbo.tab_insp WITH (NOLOCK) where num_insp = " & pobjstcItem.Num_Insp
    lSelect = lSelect & " and cod_ramo = " & pobjstcItem.Cod_Ramo

    If Not M46V999.gfObtRegistro(bdGed, lSelect, lrst, lMensagem) Then
        Call M46V999.gpGraLog(1, lMensagem)
        Exit Sub
    End If
    If lrst.EOF Then
        Call M46V999.gpMenErro("FGE0115 - Inspe��o " & pobjstcItem.Num_Insp & " inv�lida (n�o localizada).")
    Else
        If lrst("num_cnpj_cpf") <> pobjstcPed.Num_Cgc_Cpf Then
            Call M46V999.gpMenErro("FGE0115 - Inspe��o " & pobjstcItem.Num_Insp & " inv�lida (CPF/CNPJ divergente).")
        End If
    End If
    Set lrst = Nothing

End Sub

Private Sub mpConsisteEndosso(ByRef pobjstcItem As stcA46V708B)

    If pobjstcItem.Cod_DizerItem = 0 Then
        GoTo SAIDA
    End If

    '- Ass24h. N�o permitido Downgrade
    If pobjstcItem.Cod_DizerItem = e_CodDizerAlteracao Then
                If Not pobjstcItem.gStcItemApolice Is Nothing Then
                If pobjstcItem.gCodClauAss24h <> 0 And pobjstcItem.gStcItemApolice.gCodClauAss24h <> 403 Then
                    If pobjstcItem.gStcItemApolice.gCodClauAss24h > pobjstcItem.gCodClauAss24h Or pobjstcItem.gCodClauAss24h = 403 Then
                        Call M46V999.gpMenErro("FGE1095 - Ass.24h: N�o � permitido reduzir a cobertura contratada na ap�lice anterior.")
                    End If
                End If
        End If
    End If

SAIDA:
    Exit Sub

End Sub

Public Property Get M46V999() As clsM46V999
    Set M46V999 = M46V111.M46V999
End Property

Public Property Get M46O001() As clsM46O001
    Set M46O001 = M46V999.M46O001
End Property

Public Property Get M46V111() As clsM46V111
    Set M46V111 = M46V103.M46V111
End Property

Public Property Set M46V103(pClsM46V103 As clsM46V103)
    Set mClsM46V103 = pClsM46V103
End Property

Public Property Get M46V103() As clsM46V103
    If mClsM46V103 Is Nothing Then
        Err.Raise 91, Me.Class_Name, "Objeto M46V103 n�o atribu�do para a classe (" & Me.Class_Name & ")"
        '***Esta Classe DEVE utilizar uma instancia da classe clsM46V103 j�
        '   iniciada no escopo onde esta classe foi instanciada.
        'Atribua a o objeto (clsM46V103 ja instanciado) no momento que eu inicializou esta classe (atraves do Property SET)!!!
    End If

Retorno:
    Set M46V103 = mClsM46V103
End Property

Private Sub Class_Initialize()
    Set mClsM46V103 = Nothing
End Sub

Private Sub Class_Terminate()
    Set mClsM46V103 = Nothing
End Sub
