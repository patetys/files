Attribute VB_Name = "basM46V003"
Option Explicit
Private Sub mpCarregaCorretor(ByRef pCod_Corretor As Long)
    'Função: Carrega código do corretor.

    'Parâmetro de saída...pCod_Corretor...Código do corretor.

    Dim lArquivo    As String   'Endereço do arquivo.
    Dim lNumArquivo As Integer  'Utilizado com FreeFile.
    Dim lRegistro   As String   'Registro.

    '1. Inicializa variável.
    pCod_Corretor = 0

    '2. Obtém dados no arquivo P2842400.EXE.
    lArquivo = gAppPath & "P2842400.EXE"
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
    Close #lNumArquivo
End Sub
Public Function GraNovP0042100(ByVal pNosNumero As String, ByVal pNumPedido As String, _
                               ByVal pDatVigInicio As String, ByVal pDatVigTermino As String, _
                               ByVal pCodProduto As Integer, ByVal pCodProtocolo As Long, _
                               ByVal pCodSequencia As Long, ByRef pMensagem As String) As Boolean
    'Função: grava registro novo de P0042100.

    'Parâmetros de entrada...pNosNumero.......Nosso número.
    '                        pNumPedido.......Número do pedido.
    '                        pDatVigInicio....Data de vigência: início.
    '                        pDatVigTermino...Data de vigência: término.
    '                        pCodProduto......Código do produto.
    '                        pCodProtocolo....Código do protocolo.
    '                        pCodSequencia....Código da seqüência.
    'Parâmetro de saída......pMensagem........Mensagem.

    'Dim lCodCorretor    As Long         'Código do corretor.
    'Dim objP0042100     As clsA28V701A  'Objeto da classe clsA28V701A (P0042100).
    'Dim objstcP0042100  As stcA28V701B  'Objeto da classe de estrutura stcA28V701B (P0042100).
    'Dim objP0042300     As clsA28V703A  'Objeto da classe clsA28V703A (P0042300).
    'Dim objstcP0042300  As stcA28V703B  'Objeto da classe de estrutura stcA28V703B (P0042300).

    GraNovP0042100 = False

    '1. Instancio os objetos da P0042100
    'Set objstcP0042100 = New stcA28V701B
    'Set objP0042100 = New clsA28V701A

    '2. Incluo o registro na P0042100
    'If Not objP0042100.mfIncluir(objstcP0042100) Then
    '    If InStr(1, objP0042100.UltimoErro, 3022) > 0 Then
    '        pMensagem = "FTExxxx - Documento " & objstcP0042100.Nosso_Numero & _
    '                    " já existente (seguro novo)."
    '    Else
    '        pMensagem = objP0042100.UltimoErro
    '    End If
    '    GoTo Finalizar
    'End If

    'GraNovP0042100 = True

'Finalizar:
    '3. Destruo os objetos instanciados.
    'Set objP0042100 = Nothing
    'Set objstcP0042100 = Nothing
    'Set objP0042300 = Nothing
    'Set objstcP0042300 = Nothing
End Function
Private Function ImpDT1101_Item(ByVal pNumPedido As String, ByRef pNosNumero As String, _
                                ByRef pMensagem As String) As Boolean
    'Função: Importa os dados que estão na DT1101D_Item para as respectivas tabelas do Syas.

    'Parâmetros de entrada...pNumPedido...Número do pedido para encontrar o registro na tabela.
    'Parâmetros de saída.....pNosNumero...Nosso número.
    '                        pMensagem....Mensagem.

'    Dim BDSyasEmis          As ADODB.Connection         'Conexão com a base SyasEmis.
'    Dim lInsert             As String                   'Insert.
'    Dim lSelect             As String                   'Select.
'    Dim rsDT1101D_I         As ADODB.Recordset          'Registro da TAB_DT1101D_VEIC_SYAS
'    Dim rsDT1101D_C         As ADODB.Recordset          'Registro da tab_dt1101d_comum_syas
'    Dim rsDT1101D_P         As ADODB.Recordset          'Registro da tab_perfil_auto_syas
'    Dim rsDT1101D_B         As ADODB.Recordset          'Registro da tab_confirm_bonus_syas
'    Dim lRelacao            As Integer                  'Relação com o segurado (Perfil).
'    Dim objP0042200         As clsA28V702A              'Objeto da classe clsA28V702A (P0042200)
'    Dim objstcP0042200      As stcA28V702B              'Objeto da classe de estrutura stcA28V702B (P0042200)

    ImpDT1101_Item = False

    '1. Abre base de dados.
    'If gfAbrBasDados("SYAS_EMIS", gTipSYASEMIS, gEndSYASEMIS, BDSyasEmis, pMensagem) = False Then
    '    Exit Function
    'End If
    'lSelect = "SELECT * FROM TAB_DT1101D_COMUM_VVD WHERE num_prop= " & pNumPedido
    'If gfObtRegistro(BDSyasEmis, lSelect, rsDT1101D_C, pMensagem) = False Then
    '    Call gpFecha2(BDSyasEmis)
    '    Exit Function
    'End If
'    lSelect = "SELECT * FROM TAB_DT1101D_VEIC_VVD WHERE Num_prop = " & rsDT1101D_C!Num_Prop
'    If gfObtRegistro(BDSyasEmis, lSelect, rsDT1101D_I, pMensagem) = False Then
'        Call gpFecha2(BDSyasEmis)
'        Exit Function
'    End If
'
'    Do While Not rsDT1101D_I.EOF
'        lSelect = "SELECT * FROM tab_perfil_auto_vvd WHERE Num_apol = " & rsDT1101D_C!Num_Prop & _
'                  " AND num_item = " & rsDT1101D_I!Num_item
'        If gfObtRegistro(BDSyasEmis, lSelect, rsDT1101D_P, pMensagem) = False Then
'            Call gpFecha2(BDSyasEmis)
'            Exit Function
'        End If
'        lSelect = "SELECT * FROM tab_confirm_bonus_vvd WHERE Num_apol = " & rsDT1101D_C!Num_Prop & _
'                  " AND num_item = " & rsDT1101D_I!Num_item
'        If gfObtRegistro(BDSyasEmis, lSelect, rsDT1101D_B, pMensagem) = False Then
'            Call gpFecha2(BDSyasEmis)
'            Exit Function
'        End If
'        Set objstcP0042200 = New stcA28V702B
'        With objstcP0042200
'            .Nosso_Numero = pNosNumero
'            .Num_Item = rsDT1101D_I!num_item
'            .Num_Vistoria = 0
            'If Not rsDT1101D_B.EOF Then
            '    .Tip_Emissao = e_TipEmissao_RenCong
            'Else
            '    If rsDT1101D_C!cod_renov = "N" Then
            '        .Tip_Emissao = e_TipEmissao_SegNovo
            '    Else
            '        .Tip_Emissao = e_TipEmissao_RenYas
            '    End If
            'End If
'            .Cod_Iden_Vistoria = 0
'            .Sit_Item = e_SitItem_NaoCalc
'            .Lib_Cod_Produto = rsDT1101D_I!Tab_Ref_Fipe
'            .Lib_Cod_Protocolo = 0
'            .Lib_Cod_Sequencia = 0
'            .Ren_Num_Apolice = rsDT1101D_C!num_apol
'            .Ren_Num_Item = rsDT1101D_I!num_item
'            .COD_REGIAO = Left(rsDT1101D_I!Cod_Regiao, 3)
'            .CEP_Regiao = rsDT1101D_I!cod_cep_pernoite
'            .Cod_Veiculo = Format(rsDT1101D_I!cod_veic_fenaseg, "00000000")
'            If rsDT1101D_I!ano_fabr > 50 Then
'                .Ano_Fabricacao = "19" & rsDT1101D_I!ano_fabr
'            Else
'                .Ano_Fabricacao = "20" & Format(rsDT1101D_I!ano_fabr, "00")
'            End If
'            If rsDT1101D_I!ano_model_veic > 50 Then
'                .Ano_Modelo = "19" & rsDT1101D_I!ano_model_veic
'            Else
'                .Ano_Modelo = "20" & Format(rsDT1101D_I!ano_model_veic, "00")
'            End If
'            If rsDT1101D_I!num_vist_previa = "099993" Or Trim(rsDT1101D_I!num_placa) = "A/C" Then
'                .COD_IDEN_0KM = e_SIM
'            Else
'                .COD_IDEN_0KM = e_NAO
'            End If
'            If rsDT1101D_I!cod_clau_veic_novo = 19 Then
'                .Ext_0km = e_SIM
'            Else
'                .Ext_0km = e_NAO
'            End If
'            .Cod_Placa = Trim(rsDT1101D_I!num_placa)
'            .Cod_FIPE = 0
'            .Cod_Chassi = rsDT1101D_I!num_chassi
'            .Val_FIPE = 0
'            .Cod_Chassi_Estrang = rsDT1101D_I!num_chassi_estr
'            Select Case rsDT1101D_I!cod_tip_comb
'                Case "A"      'Álcool.
'                    .Tip_Combustivel = e_TipCombustivel_Alcool
'                Case "D"      'Diesel.
'                    .Tip_Combustivel = e_TipCombustivel_Diesel
'                Case "G"      'Gasolina.
'                    .Tip_Combustivel = e_TipCombustivel_Gasolina
'                Case "N"      'Outro.
'                    .Tip_Combustivel = e_TipCombustivel_Outro
'                Case Else
'                    .Tip_Combustivel = e_TipCombustivel_Gasolina
'            End Select
'            .Nom_Proprietario = rsDT1101D_I!nom_propri
'            Select Case rsDT1101D_I!cod_clau_as24h
'                Case "82"
'                    .Tip_Lim_Guincho = e_TipLimGuincho_82_Gra
'                Case "83"
'                    .Tip_Lim_Guincho = e_TipLimGuincho_83_500
'                Case "84"
'                    .Tip_Lim_Guincho = e_TipLimGuincho_84_800
'                Case "85"
'                    .Tip_Lim_Guincho = e_TipLimGuincho_85_RC
'                Case Else
'                    .Tip_Lim_Guincho = e_TipLimGuincho_82_Gra
'            End Select
'            .Val_Lim_Guincho = rsDT1101D_I!val_lim_guincho
'            Select Case rsDT1101D_I!cod_cat_tarf_auto
'                Case "000"
'                    .Tip_Cobertura = e_TipCobertura_SemFranquia
'                Case "200"
'                    .Tip_Cobertura = e_TipCobertura_IncendioRoubo
'                Case Else
'                    .Tip_Cobertura = Right(rsDT1101D_I!cod_cat_tarf_auto, 2)
'            End Select
'            If rsDT1101D_I!Per_Agrav <> "100" Then
'                .COD_IDEN_FAT_AJUSTE = e_SIM
'            Else
'                .COD_IDEN_FAT_AJUSTE = e_NAO
'            End If
'            .VAL_FAT_AJUSTE = rsDT1101D_I!Per_Agrav
'            .Tip_Produto = Right(rsDT1101D_I!cod_iden_tarf, 1)
'            Select Case Format(rsDT1101D_I!cod_clau_reser, "00")
'                Case "00"
'                    .Carro_Reserva = e_CarroReserva_SemCob
'                Case "40"
'                    .Carro_Reserva = e_CarroReserva_Cla40_7D
'                Case "41"
'                    .Carro_Reserva = e_CarroReserva_Cla41_15D
'                Case "42"
'                    .Carro_Reserva = e_CarroReserva_Cla42_30D
'            End Select
'            If rsDT1101D_P.EOF Then
'                .Cod_Iden_Questionario = e_NAO
'            Else
'                .Cod_Iden_Questionario = e_SIM
'            End If
'            If rsDT1101D_I!val_is_aces > 0 And rsDT1101D_I!cod_aces_1 = 0 Then
'                .Cod_Iden_Des_Extr = e_SIM
'            Else
'                .Cod_Iden_Des_Extr = e_NAO
'            End If
'            Select Case Trim(rsDT1101D_I!cod_desc_pickup)
'                Case 3, ""
'                    .COD_IDEN_VIDROS = e_NAO
'                Case Else
'                    .COD_IDEN_VIDROS = rsDT1101D_I!cod_desc_pickup
'            End Select
'            If rsDT1101D_I!cod_aces_1 <> 0 Then
'                .Cod_Iden_Acessorios = e_SIM
'            Else
'                .Cod_Iden_Acessorios = e_NAO
'            End If
'            .COD_ACES_1 = rsDT1101D_I!cod_aces_1
'            .VAL_IS_ACES_1 = rsDT1101D_I!val_is_aces_1
'            .COD_ACES_2 = rsDT1101D_I!cod_aces_2
'            .VAL_IS_ACES_2 = rsDT1101D_I!val_is_aces_2
'            .COD_ACES_3 = rsDT1101D_I!cod_aces_3
'            .VAL_IS_ACES_3 = rsDT1101D_I!val_is_aces_3
'            .Cod_Aces_4 = rsDT1101D_I!cod_aces_4
'            .Val_IS_Aces_4 = rsDT1101D_I!val_is_aces_4
'            .Cod_Aces_5 = rsDT1101D_I!cod_aces_5
'            .Val_IS_Aces_5 = rsDT1101D_I!val_is_aces_5
'            If rsDT1101D_I!cod_clau_1 = "53031" Or rsDT1101D_I!cod_clau_2 = "53031" Or _
'               rsDT1101D_I!cod_clau_3 = "53031" Then
'                .Soc_Dirigente = e_SIM
'            Else
'                .Soc_Dirigente = e_NAO
'            End If
'            If gfPreenchido(rsDT1101D_I!nom_benef) = True Then
'                .Cod_Iden_Beneficiario = e_SIM
'            Else
'                .Cod_Iden_Beneficiario = e_NAO
'            End If
'            .Nom_Beneficiario = rsDT1101D_I!nom_benef
'            If rsDT1101D_I!cod_aces_fabr <> "222222000" Then
'                .Ace_Fabrica = e_SIM
'            Else
'                .Ace_Fabrica = e_NAO
'            End If
'            .Cod_Iden_Ar_Condic = Mid(rsDT1101D_I!cod_aces_fabr, 1, 1)
'            .Cod_Iden_Dir_Hidra = Mid(rsDT1101D_I!cod_aces_fabr, 2, 1)
'            .Cod_Iden_Ban_Couro = Mid(rsDT1101D_I!cod_aces_fabr, 3, 1)
'            .Cod_Iden_Cam_Autom = Mid(rsDT1101D_I!cod_aces_fabr, 4, 1)
'            .Cod_Iden_Fre_ABS = Mid(rsDT1101D_I!cod_aces_fabr, 5, 1)
'            .Cod_Iden_Air_Bag = Mid(rsDT1101D_I!cod_aces_fabr, 6, 1)
'            Select Case rsDT1101D_I!cod_clau_1
'                Case "0"
'                    .Cod_Iden_Cla_Particular = e_NAO
'                Case "31070", "53031", "31010"
'                    Select Case rsDT1101D_I!cod_clau_2
'                        Case ""
'                            .Cod_Iden_Cla_Particular = e_NAO
'                        Case "31070", "53031", "31010"
'                            Select Case rsDT1101D_I!cod_clau_3
'                                Case ""
'                                    .Cod_Iden_Cla_Particular = e_NAO
'                                Case "31070", "53031", "31010"
'                                    .Cod_Iden_Cla_Particular = e_NAO
'                                Case Else
'                                    .Cod_Iden_Cla_Particular = e_SIM
'
'                            End Select
'                        Case Else
'                            .Cod_Iden_Cla_Particular = e_SIM
'                    End Select
'                Case Else
'                    .Cod_Iden_Cla_Particular = e_SIM
'            End Select
'            If Not rsDT1101D_B.EOF Then
'                .Ren_Con_Seguradora = rsDT1101D_B!cod_segura
'                .Ren_Con_Apolice = rsDT1101D_B!num_apol_segura
'                .Ren_Con_Item = rsDT1101D_B!num_item_segura
'            End If
'            .Cod_Categoria = Left(Format(rsDT1101D_I!cod_cat_tarf_auto, "000000"), 3)
'            .Ind_Sinistro = e_NAO
'            .Num_Perfil = rsDT1101D_I!dat_retirada_0km
'            If Not rsDT1101D_P.EOF Then
'                .Nom_Condutor = rsDT1101D_P!nom_motorista
'                .Cod_RG = rsDT1101D_P!num_rg
'                .Cod_CPF = rsDT1101D_P!num_cpf
'                Select Case rsDT1101D_P!cod_iden_relacao
'                    Case "P"
'                        .Cod_Iden_Relacao = e_CodRelacao_Proprio
'                    Case "F"
'                        .Cod_Iden_Relacao = e_CodRelacao_Filho
'                    Case "C"
'                        .Cod_Iden_Relacao = e_CodRelacao_Gerente
'                    Case "S"
'                        .Cod_Iden_Relacao = e_CodRelacao_Socio
'                    Case "D"
'                        .Cod_Iden_Relacao = e_CodRelacao_Diretor
'                    Case "G"
'                        .Cod_Iden_Relacao = e_CodRelacao_Gerente
'                    Case Else
'                        .Cod_Iden_Relacao = e_CodRelacao_Outro
'                End Select
'                .Dsc_Outra_Relacao = rsDT1101D_P!dsc_outros
'                .Cod_Est_Civil = rsDT1101D_P!est_civil
'                .Cod_Iden_Sexo = rsDT1101D_P!cod_iden_sexo
'                .Dat_Nasc = rsDT1101D_P!dat_nasc
'                .Dat_Pri_Habilit = Left(rsDT1101D_P!dat_habilit, 4)
'                If rsDT1101D_P!tip_cond_pess = 1 Then
'                    .Cod_Iden_Pessoas = e_SIM
'                Else
'                    .Cod_Iden_Pessoas = e_NAO
'                End If
'                .Cod_Iden_Garagem = rsDT1101D_P!cod_iden_garagem
'                .Cod_Iden_Utiliza = rsDT1101D_P!cod_util_veic
'                .Cod_Iden_Disp_Seg_1 = IIf(rsDT1101D_P!tip_disp_seg1 = "S", e_SIM, e_NAO)
'                .Cod_Iden_Disp_Seg_2 = IIf(rsDT1101D_P!tip_disp_seg2 = "S", e_SIM, e_NAO)
'                .Cod_Iden_Disp_Seg_3 = IIf(rsDT1101D_P!tip_disp_seg3 = "S", e_SIM, e_NAO)
'                .Cod_Iden_Disp_Seg_4 = IIf(rsDT1101D_P!tip_disp_seg4 = "S", e_SIM, e_NAO)
'                .Cod_Iden_Disp_Seg_5 = IIf(rsDT1101D_P!tip_disp_seg5 = "S", e_SIM, e_NAO)
'            Else
'                .Cod_Iden_Pessoas = e_NAO
'                .Cod_Iden_Garagem = "N"
'                .Cod_Iden_Utiliza = "N"
'            End If
'            .Cod_Iden_Avaria = e_NAO
'            If Not rsDT1101D_B.EOF Then
'                If rsDT1101D_B!ci_segura <> 0 Then
'                    Select Case Len(rsDT1101D_B!ci_segura)
'                        Case Is <= 5
'                            .Clausula1 = rsDT1101D_B!ci_segura
'                        Case Is <= 10
'                            .Clausula1 = Left(rsDT1101D_B!ci_segura, 5)
'                            .Clausula2 = Mid(rsDT1101D_B!ci_segura, 6)
'                        Case Else
'                            .Clausula1 = Left(rsDT1101D_B!ci_segura, 5)
'                            .Clausula2 = Mid(rsDT1101D_B!ci_segura, 6, 5)
'                            .Clausula3 = Mid(rsDT1101D_B!ci_segura, 11)
'                    End Select
'                End If
'            End If
'            .Dat_Alt = gfForData(3)
'            .Cod_User_Alt = gNomUsuario8
'            .Prazo_Iden = rsDT1101D_I!cod_iden_praz
'        End With
'        Set objP0042200 = New clsA28V702A
'        If Not objP0042200.mfIncluir(objstcP0042200) Then
'            pMensagem = objP0042200.UltimoErro
'            Call gpGraLog(1, pMensagem)
'            Call gpFecha1(rsDT1101D_I)
'            Set objP0042200 = Nothing
'            Set objstcP0042200 = Nothing
'            Exit Function
'        End If
'        If rsDT1101D_I!val_is_casco > 0 And rsDT1101D_I!num_vist_previa = "099993" Then
'            Call gfSalP0042600(1, BDSyasEmis, pNosNumero, rsDT1101D_I!num_item, gCobCasco, _
'                               Right(rsDT1101D_I!cod_iden_tarf, 1), rsDT1101D_I!val_is_casco * 0.01, _
'                               pMensagem)
'        End If
'        If rsDT1101D_I!val_is_aces > 0 Then
'            Call gfSalP0042600(1, BDSyasEmis, pNosNumero, rsDT1101D_I!num_item, gCobDesExtr, _
'                               Right(rsDT1101D_I!cod_iden_tarf, 1), rsDT1101D_I!val_is_aces * 0.01, _
'                               pMensagem)
'        End If
'        If rsDT1101D_I!val_is_rc_dm > 0 Then
'            Call gfSalP0042600(1, BDSyasEmis, pNosNumero, rsDT1101D_I!num_item, gCobRCDM, _
'                               Right(rsDT1101D_I!cod_iden_tarf, 1), rsDT1101D_I!val_is_rc_dm * 0.01, _
'                               pMensagem)
'        End If
'        If rsDT1101D_I!val_is_rc_dp > 0 Then
'            Call gfSalP0042600(1, BDSyasEmis, pNosNumero, rsDT1101D_I!num_item, gCobRCDC, _
'                               Right(rsDT1101D_I!cod_iden_tarf, 1), rsDT1101D_I!val_is_rc_dp * 0.01, _
'                               pMensagem)
'        End If
'        If rsDT1101D_I!val_is_rc_mo > 0 Then
'            Call gfSalP0042600(1, BDSyasEmis, pNosNumero, rsDT1101D_I!num_item, gCobRCMO, _
'                               Right(rsDT1101D_I!cod_iden_tarf, 1), rsDT1101D_I!val_is_rc_mo * 0.01, _
'                               pMensagem)
'        End If
'        If rsDT1101D_I!val_is_app_mort > 0 Then
'            Call gfSalP0042600(1, BDSyasEmis, pNosNumero, rsDT1101D_I!num_item, gCobAPPMorte, _
'                               Right(rsDT1101D_I!cod_iden_tarf, 1), rsDT1101D_I!val_is_app_mort * 0.01, _
'                               pMensagem)
'        End If
'        If rsDT1101D_I!val_is_app_inv > 0 Then
'            Call gfSalP0042600(1, BDSyasEmis, pNosNumero, rsDT1101D_I!num_item, gCobAPPInvalidez, _
'                               Right(rsDT1101D_I!cod_iden_tarf, 1), rsDT1101D_I!val_is_app_inv * 0.01, _
'                               pMensagem)
'        End If
'        If rsDT1101D_I!val_is_app_dmh > 0 Then
'            Call gfSalP0042600(1, BDSyasEmis, pNosNumero, rsDT1101D_I!num_item, gCobAPPDMH, _
'                               Right(rsDT1101D_I!cod_iden_tarf, 1), rsDT1101D_I!val_is_app_dmh * 0.01, _
'                               pMensagem)
'        End If
'        If rsDT1101D_I!per_bonus_auto > 0 Then
'            Call gfSalP0042700(BDSyasEmis, pNosNumero, rsDT1101D_I!num_item, gDesBonus, _
'                               rsDT1101D_I!cod_classe_bonus & " - " & rsDT1101D_I!per_bonus_auto, 0, _
'                               pMensagem)
'        End If
'        If rsDT1101D_I!per_desc_frot_rc > 0 Then
'            Call gfSalP0042700(BDSyasEmis, pNosNumero, rsDT1101D_I!num_item, gDesFrota, _
'                               rsDT1101D_I!per_desc_frot_rc, 0, pMensagem)
'        End If
'        If rsDT1101D_I!per_desc_coml_app > 0 Then
'            Call gfSalP0042700(BDSyasEmis, pNosNumero, rsDT1101D_I!num_item, gDesAPP, _
'                               rsDT1101D_I!per_desc_coml_app, 0, pMensagem)
'        End If
'        If rsDT1101D_I!per_desc_coml_auto > 0 Then
'            Call gfSalP0042700(BDSyasEmis, pNosNumero, rsDT1101D_I!num_item, gDesCasco, _
'                               rsDT1101D_I!per_desc_coml_auto, 0, pMensagem)
'        End If
'        If rsDT1101D_I!per_desc_coml_rc > 0 Then
'            Call gfSalP0042700(BDSyasEmis, pNosNumero, rsDT1101D_I!num_item, gDesRC, _
'                               rsDT1101D_I!per_desc_coml_rc, 0, pMensagem)
'        End If
'        If Right(rsDT1101D_I!per_desc_perfil, 2) > 0 Then
'            Call gfSalP0042700(BDSyasEmis, pNosNumero, rsDT1101D_I!num_item, gDesPerfil, _
'                               Right(rsDT1101D_I!per_desc_perfil, 2), 0, pMensagem)
'        End If
'        If rsDT1101D_I!per_desc_fidel > 0 Then
'            Call gfSalP0042700(BDSyasEmis, pNosNumero, rsDT1101D_I!num_item, gDesFidelidade, _
'                               rsDT1101D_I!per_desc_fidel, 0, pMensagem)
'        End If
'        rsDT1101D_I.MoveNext
'    Loop

    ImpDT1101_Item = True
End Function
Private Function ImpDT1101_Comum(ByVal pNumPedido As String, ByRef pNosNumero As String, _
                                 ByRef pMensagem As String) As Boolean
    'Função: Importa os dados que estão na DT1101D_Comum para as respectivas tabelas do Syas.

    'Parâmetro de entrada...pNumPedido...Número do pedido para encontrar o registro na tabela.
    'Parâmetros de saída....pNosNumero...Nosso número.
    '                       pMensagem....Mensagem.

'    Dim BDSyasEmis      As ADODB.Connection 'Conexão com a base SyasEmis.
'    Dim lInsert         As String           'Insert.
'    Dim lSelect         As String           'Select.
'    Dim rsDT1101D_C     As ADODB.Recordset  'Registro da DT1101D_Comum_Syas.
'    Dim rsDT1101D_I     As ADODB.Recordset  'Registro da DT1101D_Veic_Syas.
'    Dim lDat_ini_Vig    As Long             'Data do início da vigência.
'    Dim lDat_ter_vig    As Long             'Data do término da vigência.
'    Dim lCodProduto     As Integer          'Produto.
'    Dim lCodProtocolo   As Long             'Protocolo.
'    Dim lCodSequencia   As Long             'Sequencia.
'    Dim clsA99V001      As clsA99V001       'Instancia da classe.
'    Dim lDigito         As String           'Dígito do bloqueto.
'    Dim objP0042100     As clsA28V701A      'Objeto da classe clsA28V701A (P0042100).
'    Dim objstcP0042100  As stcA28V701B      'Objeto da classe de estrutura stcA28V701B (P0042100).

    ImpDT1101_Comum = False

    '1. Abre base de dados.
    'If gfAbrBasDados("SYAS_EMIS", gTipSYASEMIS, gEndSYASEMIS, BDSyasEmis, pMensagem) = False Then
    '    Exit Function
    'End If
    'lSelect = "SELECT * FROM TAB_DT1101D_COMUM_VVD WHERE num_prop = " & pNumPedido
    'If gfObtRegistro(BDSyasEmis, lSelect, rsDT1101D_C, pMensagem) = False Then
    '    Call gpFecha2(BDSyasEmis)
    '    Exit Function
    'End If
'    If Not rsDT1101D_C.EOF Then
'        If Not gfCalNosNumero(pNosNumero, pMensagem) Then
'            Call gpGraLog(2, pMensagem)
'            Call gpFecha3(rsDT1101D_C, BDSyasEmis)
'            Exit Function
'        End If
'        lSelect = "SELECT * FROM TAB_DT1101D_Veic_VVD WHERE num_prop = '" & rsDT1101D_C("Num_Prop") & "'"
'        If gfObtRegistro(BDSyasEmis, lSelect, rsDT1101D_I, pMensagem) = False Then
'            Call gpFecha2(BDSyasEmis)
'            Exit Function
'        End If
'        lDat_ini_Vig = rsDT1101D_C("dat_ini_vig")
'        lDat_ter_vig = rsDT1101D_C("dat_ter_vig")
'
'        Set objstcP0042100 = New stcA28V701B
'        With objstcP0042100
'            .Nosso_Numero = pNosNumero
'            .Cod_Ativo = "1"
'            .Tip_Documento = e_TipDocumento_Proposta
'            .Sit_Documento = e_SitDocumento_SemCalculo
'            .Sit_Processamento = e_SitProcessamento_NaoProcessado
'            .COD_SEGURADO = rsDT1101D_C("cod_segurado")
'            .NOM_SEGURADO = Trim(rsDT1101D_C("nom_segurado"))
'            If Left(rsDT1101D_C("num_cgc_cpf"), 3) = 999 Then
'                .Tip_Pessoa = e_TipPessoa_Fisica
'                .CNPJ_CPF = Mid(rsDT1101D_C("num_cgc_cpf"), 4)
'            Else
'                .Tip_Pessoa = e_TipPessoa_Juridica
'                .CNPJ_CPF = rsDT1101D_C("num_cgc_cpf")
'            End If
'            .End_Nom_Logradouro = rsDT1101D_C("nom_ender_segurado")
'            .End_Cidade = Trim(rsDT1101D_C("nom_cid_segurado"))
'            .End_UF = rsDT1101D_C("sig_uf_segurado")
'            .End_CEP = rsDT1101D_C("num_cep_segurado")
'            .End_Telefones = Trim(rsDT1101D_C("num_fone_segurado"))
'            .Cob_Diferente = e_SIM
'            .Cob_Nom_Logradouro = rsDT1101D_C("nom_ender_cobr")
'            .Cob_Cidade = Trim(rsDT1101D_C("nom_cid_cobr"))
'            .Cob_UF = rsDT1101D_C("sig_uf_cobr")
'            .Cob_CEP = rsDT1101D_C("num_cep_cobr")
'            .Cob_Telefones = Trim(rsDT1101D_C("num_fone_cobr"))
'            .Dat_Vig_Inicio = lDat_ini_Vig
'            .Dat_Vig_Termino = lDat_ter_vig
'            .Prazo_Iden = rsDT1101D_C("cod_iden_praz")
'            .Cod_Produto = Left(rsDT1101D_I("cod_iden_tarf"), 3)
'            .Cod_Protocolo = 0
'            .Cod_Sequencia = 0
'            .Qtd_Itens = rsDT1101D_C("qtd_anexo")
'            .Pag_Ant_Ind = rsDT1101D_C("cod_iden_pag_antec")
'
'            Set clsA99V001 = New clsA99V001     'DLL: calcula dígitos.
'            If clsA99V001.DIG_BLOQUETO(Format(rsDT1101D_C("num_bloq"), "00000000"), lDigito, pMensagem) = False Then
'                Call gpGraLog(2, pMensagem)
'                Exit Function
'            End If
'            .Pag_Ant_Bloqueto = rsDT1101D_C("num_bloq") & lDigito
'            If Not IsNull(rsDT1101D_C("cod_bas_cor")) Then
'                If rsDT1101D_C("cod_bas_cor") <> 0 Then
'                    .For_Pagamento = e_ForPagamento_Debito                 '
'                Else
'                    .For_Pagamento = e_ForPagamento_Carne
'                End If
'            Else
'                .For_Pagamento = e_ForPagamento_Carne
'            End If
'
'            If rsDT1101D_C("cod_bas_cor") = 0 Then
'                .Carne_Cod_Banco = rsDT1101D_C("cod_bco")
'            Else
'                .Deb_Cod_Banco = rsDT1101D_C("cod_bco")
'            End If
'            .Deb_Cod_Agencia = Left(Format(rsDT1101D_C("cod_bas_cor"), "00000000000"), 4)
'            .Deb_Cod_Conta = Right(Format(rsDT1101D_C("cod_bas_cor"), "00000000000"), 7)
'            .Ind_Ajuste = e_NAO
'            .Cor1_Codigo = rsDT1101D_C("cod_corr_1")
'            .Cor1_Cms_Casco = rsDT1101D_C("per_cms_auto_1")
'            .Cor1_Cms_RC = rsDT1101D_C("per_cms_rc_1")
'            .Cor1_Cms_APP = rsDT1101D_C("per_cms_app_1")
'            .Cor2_Codigo = rsDT1101D_C("cod_corr_2")
'            .Cor2_Cms_Casco = rsDT1101D_C("per_cms_auto_2")
'            .Cor2_Cms_RC = rsDT1101D_C("per_cms_rc_2")
'            .Cor2_Cms_APP = rsDT1101D_C("per_cms_app_2")
'            .Cor3_Codigo = rsDT1101D_C("cod_corr_3")
'            .Cor3_Cms_Casco = rsDT1101D_C("per_cms_auto_3")
'            .Cor3_Cms_RC = rsDT1101D_C("per_cms_rc_3")
'            .Cor3_Cms_APP = rsDT1101D_C("per_cms_app_3")
'            .Per_Cms_Cotacao = rsDT1101D_C("per_cms_auto_1") + _
'                            rsDT1101D_C("per_cms_auto_2") + _
'                            rsDT1101D_C("per_cms_auto_3")
'            .COD_UNID_PROD = rsDT1101D_C("cod_org") & rsDT1101D_C("cod_escr")
'            .Cod_Produtor = rsDT1101D_C("cod_prod")
'            If rsDT1101D_C("cod_coss_1") > 0 Then
'                .Ind_Cosseguro = e_SIM
'            Else
'                .Ind_Cosseguro = e_NAO
'            End If
'            .Cos1_Codigo = rsDT1101D_C("cod_coss_1")
'            .Cos1_Percentual = rsDT1101D_C("per_coss_1")
'            .Cos2_Codigo = rsDT1101D_C("cod_coss_2")
'            .Cos2_Percentual = rsDT1101D_C("per_coss_2")
'            .Cos3_Codigo = rsDT1101D_C("cod_coss_3")
'            .Cos3_Percentual = rsDT1101D_C("per_coss_3")
'            .Cos4_Codigo = rsDT1101D_C("cod_coss_4")
'            .Cos4_Percentual = rsDT1101D_C("per_coss_4")
'            .Cos5_Codigo = rsDT1101D_C("cod_coss_5")
'            .Cos5_Percentual = rsDT1101D_C("per_coss_5")
'            .Cus_Emissao = rsDT1101D_C("val_cust_emis")
'            .Cus_Emi_Parcelado = rsDT1101D_C("cod_tip_cust_emis")
'            .Cod_Registro = rsDT1101D_C("cod_reg")
'            .Qtd_Parcelas_VM = rsDT1101D_C("cod_form_pag")
'            .Ven_Parcela1_VM = rsDT1101D_C("dat_venc") & ","
'            If rsDT1101D_C("cod_form_pag") = 1 Then
'                .Jur_Ant_01 = rsDT1101D_C("per_juro_frac")
'            Else
'                .Jur_Ant_01 = 3
'            End If
'            If rsDT1101D_C("cod_form_pag") = 2 Then
'                .Jur_Ant_02 = rsDT1101D_C("per_juro_frac")
'            Else
'                .Jur_Ant_02 = 3
'            End If
'            If rsDT1101D_C("cod_form_pag") = 3 Then
'                .Jur_Ant_03 = rsDT1101D_C("per_juro_frac")
'            Else
'                .Jur_Ant_03 = 3
'            End If
'            If rsDT1101D_C("cod_form_pag") = 4 Then
'                .Jur_Ant_04 = rsDT1101D_C("per_juro_frac")
'            Else
'                .Jur_Ant_04 = 3
'            End If
'            If rsDT1101D_C("cod_form_pag") = 5 Then
'                .Jur_Ant_05 = rsDT1101D_C("per_juro_frac")
'            Else
'                .Jur_Ant_05 = 3
'            End If
'            If rsDT1101D_C("cod_form_pag") = 6 Then
'                .Jur_Ant_06 = rsDT1101D_C("per_juro_frac")
'            Else
'                .Jur_Ant_06 = 3
'            End If
'            If rsDT1101D_C("cod_form_pag") = 7 Then
'                .Jur_Ant_07 = rsDT1101D_C("per_juro_frac")
'            Else
'                .Jur_Ant_07 = 3
'            End If
'            If rsDT1101D_C("cod_form_pag") = 8 Then
'                .Jur_Ant_08 = rsDT1101D_C("per_juro_frac")
'            Else
'                .Jur_Ant_08 = 3
'            End If
'            If rsDT1101D_C("cod_form_pag") = 9 Then
'                .Jur_Ant_09 = rsDT1101D_C("per_juro_frac")
'            Else
'                .Jur_Ant_09 = 3
'            End If
'            If rsDT1101D_C("cod_form_pag") = 10 Then
'                .Jur_Ant_10 = rsDT1101D_C("per_juro_frac")
'            Else
'                .Jur_Ant_10 = 3
'            End If
'            If rsDT1101D_C("cod_form_pag") = 11 Then
'                .Jur_Ant_11 = rsDT1101D_C("per_juro_frac")
'            Else
'                .Jur_Ant_11 = 3
'            End If
'            If rsDT1101D_C("cod_form_pag") = 12 Then
'                .Jur_Ant_12 = rsDT1101D_C("per_juro_frac")
'            Else
'                .Jur_Ant_12 = 3
'            End If
'            If rsDT1101D_C("cod_form_pag") = 1 Then
'                .Jur_Pos_01 = rsDT1101D_C("per_juro_frac")
'            Else
'                .Jur_Pos_01 = 3
'            End If
'            If rsDT1101D_C("cod_form_pag") = 2 Then
'                .Jur_Pos_02 = rsDT1101D_C("per_juro_frac")
'            Else
'                .Jur_Pos_02 = 3
'            End If
'            If rsDT1101D_C("cod_form_pag") = 3 Then
'                .Jur_Pos_03 = rsDT1101D_C("per_juro_frac")
'            Else
'                .Jur_Pos_03 = 3
'            End If
'            If rsDT1101D_C("cod_form_pag") = 4 Then
'                .Jur_Pos_04 = rsDT1101D_C("per_juro_frac")
'            Else
'                .Jur_Pos_04 = 3
'            End If
'            If rsDT1101D_C("cod_form_pag") = 5 Then
'                .Jur_Pos_05 = rsDT1101D_C("per_juro_frac")
'            Else
'                .Jur_Pos_05 = 3
'            End If
'            If rsDT1101D_C("cod_form_pag") = 6 Then
'                .Jur_Pos_06 = rsDT1101D_C("per_juro_frac")
'            Else
'                .Jur_Pos_06 = 3
'            End If
'            If rsDT1101D_C("cod_form_pag") = 7 Then
'                .Jur_Pos_07 = rsDT1101D_C("per_juro_frac")
'            Else
'                .Jur_Pos_07 = 3
'            End If
'            If rsDT1101D_C("cod_form_pag") = 8 Then
'                .Jur_Pos_08 = rsDT1101D_C("per_juro_frac")
'            Else
'                .Jur_Pos_08 = 3
'            End If
'            If rsDT1101D_C("cod_form_pag") = 9 Then
'                .Jur_Pos_09 = rsDT1101D_C("per_juro_frac")
'            Else
'                .Jur_Pos_09 = 3
'            End If
'            If rsDT1101D_C("cod_form_pag") = 10 Then
'                .Jur_Pos_10 = rsDT1101D_C("per_juro_frac")
'            Else
'                .Jur_Pos_10 = 3
'            End If
'            If rsDT1101D_C("cod_form_pag") = 11 Then
'                .Jur_Pos_11 = rsDT1101D_C("per_juro_frac")
'            Else
'                .Jur_Pos_11 = 3
'            End If
'            If rsDT1101D_C("cod_form_pag") = 12 Then
'                .Jur_Pos_12 = rsDT1101D_C("per_juro_frac")
'            Else
'                .Jur_Pos_12 = 3
'            End If
'            .Dat_Alt = gfForData(3)
'            .Cod_User_Alt = gNomUsuario8
'        End With
'
'        Set objP0042100 = New clsA28V701A
'        If Not objP0042100.mfIncluir(objstcP0042100) Then
'            pMensagem = objP0042100.UltimoErro
'            Call gpFecha1(rsDT1101D_C)
'            Exit Function
'        End If
'    Else
'        Call gpGraLog(2, "Nenhuma apólice foi localizada")
'        ImpDt1101_Comum = False
'        Exit Function
'    End If
'
'    ImpDt1101_Comum = True
End Function
