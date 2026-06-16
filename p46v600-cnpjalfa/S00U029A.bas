Attribute VB_Name = "xxxS00U029"
Option Explicit
Public gS00U029_Bonus           As Double   'Percentual de bônus..................PER-BONUS-S00U029.
Public gS00U029_BonResseguro    As Double   'Percentual de bônus para resseguro...PER-BONUS-RESS-S00U029.
Public gS00U029_Produto         As String   'Código do produto....................COD-PRODUTO-S00U029.
Public gS00U029_Ramo            As String   'Código do ramo.......................COD-RAMO-S00U029.
Public gS00U029_Renovacao       As String   'Renovação Yasuda/Congênere...........COD-RENOV-S00U029.
Public Function gfS00U029() As Boolean
    'Saída...True....Ok.
    '        False...houve erro.

    
    gfS00U029 = False
    gS00U029_BonResseguro = 0

    Select Case gS00U029_Ramo
        Case "112"
            Select Case gS00U029_Produto
                Case "0003", "1010"
                    Select Case gS00U029_Renovacao
                        Case "Y", "C"
                            Select Case gS00U029_Bonus
                                Case 0, 3
                                    gfS00U029 = True
                                    gS00U029_BonResseguro = 0
                                Case 6
                                    gfS00U029 = True
                                    gS00U029_BonResseguro = 5
                                Case 8
                                    gfS00U029 = True
                                    gS00U029_BonResseguro = 10
                                Case 10
                                    gfS00U029 = True
                                    gS00U029_BonResseguro = 15
                                Case 15
                                    gfS00U029 = True
                                    gS00U029_BonResseguro = 20
                            End Select
                    End Select
                Case "1020", "1030", "1040"
                    Select Case gS00U029_Renovacao
                        Case "Y", "C"
                            Select Case gS00U029_Bonus
                                Case 0
                                    gfS00U029 = True
                                    gS00U029_BonResseguro = 0
                                Case 3
                                    gfS00U029 = True
                                    gS00U029_BonResseguro = 3
                                Case 6
                                    gfS00U029 = True
                                    gS00U029_BonResseguro = 6
                                Case 8
                                    gfS00U029 = True
                                    gS00U029_BonResseguro = 8
                                Case 10
                                    gfS00U029 = True
                                    gS00U029_BonResseguro = 10
                                Case 15
                                    gfS00U029 = True
                                    gS00U029_BonResseguro = 15
                            End Select
                    End Select
            End Select
        Case "113"
            Select Case gS00U029_Produto
                Case "0002"
                    Select Case gS00U029_Renovacao
                        Case "Y"
                            Select Case gS00U029_Bonus
                                Case 0
                                    gfS00U029 = True
                                    gS00U029_BonResseguro = 0
                                Case 5
                                    gfS00U029 = True
                                    gS00U029_BonResseguro = 5
                                Case 10
                                    gfS00U029 = True
                                    gS00U029_BonResseguro = 10
                                Case 15
                                    gfS00U029 = True
                                    gS00U029_BonResseguro = 15
                                Case 20
                                    gfS00U029 = True
                                    gS00U029_BonResseguro = 20
                                Case 25
                                    gfS00U029 = True
                                    gS00U029_BonResseguro = 25
                            End Select
                        Case "C"
                            Select Case gS00U029_Bonus
                                Case 0, 3, 7, 12, 15, 18
                                    gfS00U029 = True
                                    gS00U029_BonResseguro = 0
                            End Select
                    End Select
                Case "1001", "1002", "1003", "1004", "1005", "1006", "1007", "1008", "1009", "1010"
                    If gS00U029_Renovacao = "Y" And gS00U029_Bonus = 1200 Then
                        gfS00U029 = True
                        gS00U029_BonResseguro = 0
                    End If
            End Select
        Case "114"
            Select Case gS00U029_Produto
                Case "0001"
                    Select Case gS00U029_Renovacao
                        Case "Y", "C"
                            Select Case gS00U029_Bonus
                                Case 0
                                    gfS00U029 = True
                                    gS00U029_BonResseguro = 0
                                Case 10
                                    gfS00U029 = True
                                    gS00U029_BonResseguro = 5
                                Case 15
                                    gfS00U029 = True
                                    gS00U029_BonResseguro = 10
                                Case 20
                                    gfS00U029 = True
                                    gS00U029_BonResseguro = 15
                                Case 25
                                    gfS00U029 = True
                                    gS00U029_BonResseguro = 20
                                Case 30
                                    gfS00U029 = True
                                    gS00U029_BonResseguro = 25
                            End Select
                    End Select
                Case "0002"
                    Select Case gS00U029_Renovacao
                        Case "Y"
                            Select Case gS00U029_Bonus
                                Case 0
                                    gfS00U029 = True
                                    gS00U029_BonResseguro = 0
                                Case 10
                                    gfS00U029 = True
                                    gS00U029_BonResseguro = 5
                                Case 15
                                    gfS00U029 = True
                                    gS00U029_BonResseguro = 10
                                Case 20
                                    gfS00U029 = True
                                    gS00U029_BonResseguro = 15
                                Case 25
                                    gfS00U029 = True
                                    gS00U029_BonResseguro = 20
                                Case 30
                                    gfS00U029 = True
                                    gS00U029_BonResseguro = 25
                            End Select
                        Case "C"
                            Select Case gS00U029_Bonus
                                Case 0, 7, 12, 15, 18, 20
                                    gfS00U029 = True
                                    gS00U029_BonResseguro = 0
                            End Select
                    End Select
            End Select
    End Select
End Function
