VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V703Brpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mStcA46V703B As stcA46V703B

Public Property Set stcOrigem(ByRef pStcA46V703B As stcA46V703B)
    Set mStcA46V703B = pStcA46V703B
End Property

Public Property Get CodClausula() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    If mStcA46V703B.Cod_Clau <> 403 Then
    lretorno = mStcA46V703B.Cod_Clau
    End If
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    CodClausula = lretorno
End Property

Public Property Get DescClausula() As String
    Dim lretorno As String
    Dim lAuxiliar As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lAuxiliar = IIf(mStcA46V703B.gEhAss24h, "Assistência 24 horas: ", "")
    lretorno = lAuxiliar & mStcA46V703B.gDescClausula
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    lretorno = IIf(lretorno = "", " ", lretorno)
    DescClausula = lretorno
End Property

Public Property Get TextoClausula() As String
    Dim lretorno As String
    '    Select Case tipoRelatorio
    '        Case eTipCalcRPT_CotacIndividual
    lretorno = " "
    
    Dim lNomeSeguradora As String

    ' Obter o nome da Seguradora
    lNomeSeguradora = "Yasuda Maritima"
    If mStcA46V703B.M46V999.gDataCorteSompo <> "" Then
        If IsNumeric(mStcA46V703B.M46V999.gDataCorteSompo) Then
            If CLng(mStcA46V703B.M46V999.gDataCorteSompo) <= mStcA46V703B.M46V999.gHoje Then
                lNomeSeguradora = "Sompo"
            End If
        End If
    End If
    
    If mStcA46V703B.Cod_Ramo = e_CodRamo_Empresarial Then
        Select Case mStcA46V703B.Cod_Clau
        Case 318
            If mStcA46V703B.ObjPai.Cod_Iden_Tarif >= 1016 Then  '1100705 - Tarifa 1016 e 1017 Agosto 2011
                lretorno = "Fica entendido e acordado que, além dos riscos excluídos previstos na cobertura de Pátio, sub-item 4.2.30 e na " & _
                           "Cláusula 5ª - Riscos Excluídos das Condições Gerais , esta cobertura não garantirá os veículos estacionados em " & _
                           "locais abertos, cujas as chaves originais e reservas não estiverem guardadas, após o horário comercial, dentro de " & _
                           "armários, gavetas ou claviculários trancados a chave. Entende-se por após o horário comercial, para fins desta " & _
                           "cláusula, o período compreendido entre 20:00 hs e 08:00 hs e horário comercial aquele compreendido entre 08:00 hs e 20:00 hs."
            Else
                lretorno = "Fica entendido e acordado que além dos riscos excluídos previstos na cobertura " & _
                           "de Pátio, sub-item 4.2.30, esta cobertura não garante perdas e danos conseqüentes, direta ou " & _
                           "indiretamente, de: - subtração de veículos de propriedade exclusiva do segurado, novos e/ou usados, " & _
                           "que compõem o seu estoque, destinado à venda exposição e/ou revenda, cujas as chaves originais e reservas " & _
                           "não estiverem trancadas em cofre-forte de aço com peso mínimo de 50 quilos, com chave e segredo acionados. " & _
                           "Esta exclusão não se aplica quando a concessionária estiver em funcionamento em seu horário comercial, devendo " & _
                           "este horário ser compreendido entre 08:00 hs e 20:00 hs. Entre as 20:00 hs e 08:00 hs, as chaves dos veículos " & _
                           "deverão permanecer dentro dos cofres conforme condição descrita nesta clausula."
            End If
        Case 319
            lretorno = "Fica entendido e acordado que além dos riscos excluídos previstos na cobertura " & _
                       "de Responsabilidade Civil Concessionária, esta cobertura não garante perdas e danos " & _
                       "conseqüentes, direta ou indiretamente, de: - subtração de veículos de propriedade de " & _
                       "terceiros, novos e/ou usados, que estejam em poder do segurado para conserto e /ou " & _
                       "manutenção, ou em consignação para fins de exposição ou revenda, cujas as chaves originais " & _
                       "e reservas não estiverem trancadas em cofre-forte de aço com peso mínimo de 50 quilos, com " & _
                       "chave e segredo acionados. Esta exclusão não se aplica quando a concessionária estiver em " & _
                       "funcionamento em seu horário comercial, devendo este ser horário compreendido entre 08:00 hs e 20:00 hs. " & _
                       "Entre as 20:00 hs e 08:00 hs, as chaves dos veículos deverão permanecer dentro dos cofres conforme condição descrita nesta clausula."
        Case 320
            If mStcA46V703B.ObjPai.Cod_Iden_Tarif >= 1018 Then    '1100864 - Tarifa 1018 e 1019 Setembro 2011
                lretorno = "Fica entendido e acordado que ao contrário do disposto na alínea 'a', da Cláusula 6ª - Bens não compreendidos no seguro das " & _
                           "Condições Gerais do produto " & lNomeSeguradora & " Empresarial, o presente seguro garantirá, nos termos das coberturas contratadas na apólice, os bens dos hóspedes do hotel " & _
                           "segurado até o limite de R$ 20.000,00. Fica entendido e acordado ainda que, a exceção descrita nesta cláusula não se aplicará às coberturas de Responsabilidade Civil, " & _
                           "aos bens que sejam excluídos do seguro e às exclusões constantes das coberturas contratadas."
            End If
        Case 321
            If mStcA46V703B.ObjPai.Cod_Iden_Tarif >= 1018 Then    '1100864 - Tarifa 1018 e 1019 Setembro 2011
                lretorno = "Fica entendido e acordado que ao contrário do que possa constar na Cláusula 6ª - Bens não compreendidos no seguro das Condições Gerais do produto " & lNomeSeguradora & " Empresarial, " & _
                           "a cobertura acessória de Roubo de Bens garantirá os bens descritos nas alíneas 'a', 'b' e 'c' abaixo, desde que sejam pertencentes aos hóspedes do hotel segurado: " & _
                           "a) raridades, antigüidades, jóias, pedras, metais preciosos, relógios, quadros e objetos de arte, desde que dentro de cofres-fortes devidamente fechados com chave e segredo e que tenham " & _
                           "sido confiados ao Segurado mediante documento de recepção protocolado e assinado pelo respectivo hóspede. Não estarão cobertos os bens deixados nos cofres " & _
                           "existentes no interior dos aposentos dos hóspedes. " & _
                           "b) Objetos de uso pessoal, exceto perfumes, cosméticos, bebidas, medicamentos. " & _
                           "c) Livros, telefones celulares e seus acessórios, bem como aparelhos de transmissão e recepção de rádio freqüência. " & _
                           "Tudo de conformidade com os riscos cobertos da cobertura de Roubo de Bens. " & _
                           "Esta cobertura está limitada a R$ 20.000,00 ou o LMG contratado para a cobertura acessória de Roubo de Bens, o que for menor. "
            End If
        Case 322
            If mStcA46V703B.ObjPai.Cod_Iden_Tarif >= 1018 Then    '1100864 - Tarifa 1018 e 1019 Setembro 2011
                lretorno = "Fica entendido e acordado que estarão amparados pela cobertura de Roubo de Valores no Interior do Estabelecimento, sob a forma de reembolso ao Segurado, os valores pertencentes aos " & _
                           "hóspedes enquanto hospedados no hotel segurado e localizados no interior do hotel segurado, desde que dentro de cofres-fortes devidamente fechados com chave e segredo e que tenham sido confiados ao Segurado " & _
                           "mediante documento de recepção protocolado e assinado pelo respectivo hóspede. Tudo de conformidade com os riscos cobertos da cobertura de Roubo de Valores no Interior do Estabelecimento. " & _
                           "Esta cobertura está limitada a R$ 20.000,00 ou o LMG contratado para a cobertura acessória de Roubo de Valores no Interior do Estabelecimento, o que for menor."
            End If
        End Select
    End If
    '        Case eTipCalcRPT_CotacColetiva
    '        Case eTipCalcRPT_PropIndividual
    '        Case eTipCalcRPT_PropColetiva
    '    End Select
    TextoClausula = lretorno
End Property

'Public Property Get XXXXX() As String
'Dim lRetorno As String
''    Select Case tipoRelatorio
''        Case eTipCalcRPT_CotacIndividual
'            lRetorno = mStcA46V703B.WWWWW
''        Case eTipCalcRPT_CotacColetiva
''        Case eTipCalcRPT_PropIndividual
''        Case eTipCalcRPT_PropColetiva
''    End Select
'    lRetorno = IIf(lRetorno = "", " ", lRetorno)
'    XXXXX = lRetorno
'End Property


Private Sub Class_Initialize()
    Set mStcA46V703B = Nothing
End Sub

Private Sub Class_Terminate()
    Set mStcA46V703B = Nothing
End Sub
