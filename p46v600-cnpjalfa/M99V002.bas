Attribute VB_Name = "basM99V002"
'ANTIGO STSPRINT.BAS.

Option Explicit
'Todas as medidas devem ser em centímetros.
'Cores permitidas para fonte: "PRETO", "AZUL", "VERDE", "VERMELHO", "BRANCO" e "AMARELO".
'Papel permitido:
'   DEFAULT SIZE
'   LETTER 8 1/2 X 11 IN
'   LETTER SMALL 8 1/2 X 11 IN
'   TABLOID 11 X 17 IN
'   LEDGER 17 X 11 IN
'   LEGAL 8 1/2 X 14 IN
'   STATEMENT 5 1/2 X 8 1/2 IN
'   EXECUTIVE7 1/2 X 10 IN
'   A3 297 X 420 MM
'   A4 210 X 297 MM
'   A4 SMALL 210 X 297 MM
'   A5 148 X 210 MM
'   B4 250 X 354 MM
'   B5 182 X 257 MM
'   FOLIO 8 1/2 X 13 IN
'   QUARTO 215 X 275 MM
'   10X14 IN
'   11X17 IN
'   NOTE 8 1/2 X 11 IN
'   ENVELOPE #9 3 7/8 X 8 7/8
'   ENVELOPE #10 4 1/8 X 9 1/2
'   ENVELOPE #11 4 1/2 X 10 3/8
'   ENVELOPE #12 4 ¾ X 11
'   ENVELOPE #12 4 ¾ X 11
'   C SIZE SHEET
'   D SIZE SHEET
'   E SIZE SHEET
'   CUSTOM SIZE
'   Em "Custom" devem ser fornecidas altura e largura do papel.
'   Para "A4", "Letter", "Legal" e demais, altura e largura devem ser zero (0).
'Opção de impressão em print: 0 = Preview; 1 = Print e 2 = arquivo imagem.
'Orientação: 0 = portrait e 1 = land.

Declare Sub clearDados Lib "PrintSts.DLL" ()

Declare Sub setDados Lib "PrintSts.DLL" (ByVal Index As Integer, _
                                         ByVal campo As String, _
                                         ByVal posx As Double, _
                                         ByVal posy As Double, _
                                         ByVal nomeFonte As String, _
                                         ByVal tamFonte As Integer, _
                                         ByVal color As String, _
                                         ByVal bold As String, _
                                         ByVal italico As String, _
                                         ByVal sublinhado As String, _
                                         ByVal riscado As String)

Declare Sub setPaper Lib "PrintSts.DLL" (ByVal paper As String, _
                                         ByVal lenght As Double, _
                                         ByVal Width As Double)

Declare Sub setImage Lib "PrintSts.DLL" (ByVal imagem As String, _
                                         ByVal Top As Double, _
                                         ByVal Left As Double, _
                                         ByVal Height As Double, _
                                         ByVal Width As Double)

Declare Sub setArqImage Lib "PrintSts.DLL" (ByVal arq As String)

Declare Sub PrintRel Lib "PrintSts.DLL" (ByVal opcao As String, _
                                         ByVal orientacao As String)

Declare Sub setArqConfig Lib "PrintSts.DLL" (ByVal arq As String, _
                                             ByVal opcaoArq As String)
'Obs.: opcaoArq: "0" = binário e "1" = texto.

Declare Sub setArqDados Lib "PrintSts.DLL" (ByVal arq As String)

Declare Sub PrintReld Lib "PrintSts.DLL" (ByVal opcao As String)
