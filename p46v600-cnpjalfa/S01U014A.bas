Attribute VB_Name = "xxxS01U014"
Option Explicit
Public gS01U014_Ramo        As String       'Código de ramo........COD-RAMO-S01U014.
Public gS01U014_Produto     As String       'Código de produto.....COD-PRODUTO-S01U014.
Public gS01U014_Susep       As String       'Número de inscrição...NUM-PROC-SUSEP-S01U014.
Public gS01U014_Aceita      As String       'Código de aceitação...COD-ACEITA-S01U014.
Public gS01U014_Mensagem    As String       'Mensagem..............MSG-S01U014.
Public gS01U014_IniVigencia As String       'Início da vigência....DAT-INI-VIG-S01U014.
Public gS01U014_Retorno     As Byte         'Código de retorno.
Public gS01U014_TerVigencia As String       'Término da vigência...DAT-TER-VIG-S01U014.
Public gS01U014_TabTarifa   As String       'Tabela de tarifa......TAB-TAR-S01U014.

Private mTabSusep()         As String       'Tabela da Susep.
Public Sub gpS01U014()
    Dim li  As Byte     'Utilizado com For...Next.

    '1. Verifica parâmetros (001-00-VERIFICA-PARM).
    If Not IsNumeric(gS01U014_Ramo) And Not IsNumeric(gS01U014_Produto) Then
        gS01U014_Retorno = 1
        gS01U014_Mensagem = "Códigos de ramo e produto não numéricos."
        gS01U014_Susep = ""
        gS01U014_Aceita = ""
        gS01U014_IniVigencia = 0
        gS01U014_TerVigencia = 0
        gS01U014_TabTarifa = 0
        Exit Sub
    End If
    If Not IsNumeric(gS01U014_Ramo) Then
        gS01U014_Retorno = 2
        gS01U014_Mensagem = "Código de ramo não numérico."
        gS01U014_Susep = ""
        gS01U014_Aceita = ""
        gS01U014_IniVigencia = 0
        gS01U014_TerVigencia = 0
        gS01U014_TabTarifa = 0
        Exit Sub
    End If
    If Not IsNumeric(gS01U014_Produto) Then
        gS01U014_Retorno = 3
        gS01U014_Mensagem = "Código de produto não numérico."
        gS01U014_Susep = ""
        gS01U014_Aceita = ""
        gS01U014_IniVigencia = 0
        gS01U014_TerVigencia = 0
        gS01U014_TabTarifa = 0
        Exit Sub
    End If

    '2. Processa (007-00-PROCESSA).
    Call mpCarTabSusep
    For li = 1 To UBound(mTabSusep)
        If gS01U014_Ramo = Mid(mTabSusep, 1, 3) Then
            gS01U014_Retorno = 0
            gS01U014_Mensagem = ""
            gS01U014_Susep = Mid(mTabSusep, 15, 30)
            gS01U014_Aceita = Mid(mTabSusep, 46, 1)
            gS01U014_IniVigencia = Mid(mTabSusep, 48, 8)
            gS01U014_TerVigencia = Mid(mTabSusep, 57, 8)
            gS01U014_TabTarifa = Mid(mTabSusep, 66, 4)
            Exit Sub
        End If
    Next li
    gS01U014_Retorno = 4
    gS01U014_Mensagem = "Código de ramo e/ou produto não localizado."
    gS01U014_Susep = "***************"
    gS01U014_Aceita = " "
    gS01U014_IniVigencia = 0
    gS01U014_TerVigencia = 0
    gS01U014_TabTarifa = 0
End Sub
Private Sub mpCarTabSusep()
    'Procedure: carrega tabela da SUSEP.

    'Formato da tabela:
    'Posições de 01 a 03 (03 posições)...Código do ramo........NUM-SUSEP-RAMO-WS.
    'Posições de 04 a 13 (10 posições)...Código do produto.....NUM-SUSEP-PROD-WS.
    'Posição     14......(01 posição)....Traço.................FILLER.
    'Posições de 15 a 44 (31 posições)...Código Susep..........NUM-SUSEP-CODIGO-WS.
    'Posição     45......(01 posição)....Traço.................FILLER.
    'Posição     46......(01 posição)....Código de aceitação...NUM-SUSEP-ACEITA-WS.
    'Posição     47......(01 posição)....Traço.................FILLER.
    'Posições de 48 a 55 (08 posições)...Início de vigência....DAT-INI-VIG-WS.
    'Posição     56......(01 posição)....Traço.................FILLER.
    'Posições de 57 a 64 (08 posições)...Término de vigência...DAT-TER-VIG-WS.
    'Posição     65......(01 posição)....Traço.................FILLER.
    'Posições de 66 a 69 (04 posições)...Tabela de tarifa......TAB-TAR-WS.

    'Atenção: carregar tabela em ordem crescente de ramo e produto.

    Dim li  As Byte     'Índice.
    ReDim mTabSusep(61)

    li = 0
    li = li + 1
    mTabSusep(li) = "1100000000000-********************          -S-00000000-99999999-0001"
    li = li + 1
    mTabSusep(li) = "1120000000000-15414.006484/98-93            -S-00000000-99999999-0003"
    li = li + 1
    mTabSusep(li) = "1120000001010-DEFINIR QUANDO IMPLANTAR      -S-00000000-00000000-1010"
    li = li + 1
    mTabSusep(li) = "1130000000000-001.3749/94                   -S-00000000-19990630-0001"
    li = li + 1
    mTabSusep(li) = "1130000001099-10005040/99-04                -S-19990701-99999999-0002"
    li = li + 1
    mTabSusep(li) = "1140000000000-001.5536/94                   -S-00000000-99999999-0001"
    li = li + 1
    mTabSusep(li) = "1200000000000-********************          -S-00000000-99999999-0001"
    li = li + 1
    mTabSusep(li) = "1300000000000-********************          -S-00000000-99999999-0001"
    li = li + 1
    mTabSusep(li) = "1500000000000-********************          -S-00000000-99999999-0001"
    li = li + 1
    mTabSusep(li) = "1700000000000-********************          -S-00000000-99999999-0001"
    li = li + 1
    mTabSusep(li) = "2110000000000-********************          -S-00000000-99999999-0001"
    li = li + 1
    mTabSusep(li) = "2120000000000-********************          -S-00000000-99999999-0001"
    li = li + 1
    mTabSusep(li) = "3100000000000-001.5306/96                   -S-00000000-99999999-0001"
    li = li + 1
    mTabSusep(li) = "3100000000799-001.5306/96                   -S-00000000-99999999-0006"
    li = li + 1
    mTabSusep(li) = "3100000001299-001.5306/96                   -S-00000000-99999999-0007"
    li = li + 1
    mTabSusep(li) = "3100000007001-001.5306/96                   -S-20000601-99999999-0008"
    li = li + 1
    mTabSusep(li) = "3100000007002-001.5306/96                   -S-20000601-99999999-0008"
    li = li + 1
    mTabSusep(li) = "3100000008201-001.5306/96                   -S-20000601-99999999-0008"
    li = li + 1
    mTabSusep(li) = "3100000008202-001.5306/96                   -S-20000601-99999999-0008"
    li = li + 1
    mTabSusep(li) = "3100000008211-001.5306/96                   -S-20000701-99999999-0008"
    li = li + 1
    mTabSusep(li) = "3100000008212-001.5306/96                   -S-20000701-99999999-0008"
    li = li + 1
    mTabSusep(li) = "3100000008221-001.5306/96                   -S-20000701-99999999-0008"
    li = li + 1
    mTabSusep(li) = "3100000008222-001.5306/96                   -S-20000701-99999999-0008"
    li = li + 1
    mTabSusep(li) = "3300000000000-********************          -S-00000000-99999999-0001"
    li = li + 1
    mTabSusep(li) = "3310000000000-001.002803/95                 -S-00000000-99999999-0001"
    li = li + 1
    mTabSusep(li) = "3400000000000-********************          -S-00000000-99999999-0001"
    li = li + 1
    mTabSusep(li) = "3500000000000-********************          -S-00000000-99999999-0001"
    li = li + 1
    mTabSusep(li) = "4100000000000-********************          -S-00000000-99999999-0001"
    li = li + 1
    mTabSusep(li) = "4300000000000-********************          -S-00000000-99999999-0001"
    li = li + 1
    mTabSusep(li) = "4400000000000-********************          -S-00000000-99999999-0001"
    li = li + 1
    mTabSusep(li) = "4600000000000-********************          -S-00000000-99999999-0001"
    li = li + 1
    mTabSusep(li) = "4800000000000-********************          -S-00000000-99999999-0001"
    li = li + 1
    mTabSusep(li) = "4900000000000-********************          -S-00000000-99999999-0001"
    li = li + 1
    mTabSusep(li) = "5100000000000-********************          -S-00000000-99999999-0001"
    li = li + 1
    mTabSusep(li) = "5300000000000-********************          -S-00000000-99999999-0001"
    li = li + 1
    mTabSusep(li) = "5400000000000-********************          -S-00000000-99999999-0001"
    li = li + 1
    mTabSusep(li) = "5500000000000-********************          -S-00000000-99999999-0001"
    li = li + 1
    mTabSusep(li) = "5600000000000-********************          -S-00000000-99999999-0001"
    li = li + 1
    mTabSusep(li) = "5700000000000-********************          -S-00000000-99999999-0001"
    li = li + 1
    mTabSusep(li) = "6200000000000-********************          -S-00000000-99999999-0001"
    li = li + 1
    mTabSusep(li) = "6400000000000-********************          -S-00000000-99999999-0001"
    li = li + 1
    mTabSusep(li) = "6600000000000-********************          -S-00000000-99999999-0001"
    li = li + 1
    mTabSusep(li) = "6700000000000-********************          -S-00000000-99999999-0001"
    li = li + 1
    mTabSusep(li) = "6800000000000-********************          -S-00000000-99999999-0001"
    li = li + 1
    mTabSusep(li) = "7100000000000-********************          -S-00000000-99999999-0001"
    li = li + 1
    mTabSusep(li) = "7110000000000-********************          -S-00000000-99999999-0001"
    li = li + 1
    mTabSusep(li) = "7120000000000-********************          -S-00000000-99999999-0001"
    li = li + 1
    mTabSusep(li) = "7130000000000-********************          -S-00000000-99999999-0001"
    li = li + 1
    mTabSusep(li) = "7140000000000-********************          -S-00000000-99999999-0001"
    li = li + 1
    mTabSusep(li) = "7300000000000-********************          -S-00000000-99999999-0001"
    li = li + 1
    mTabSusep(li) = "7500000000000-********************          -S-00000000-99999999-0001"
    li = li + 1
    mTabSusep(li) = "8100000008101-15414.003268/97-51            -S-00000000-99999999-0001"
    li = li + 1
    mTabSusep(li) = "8100000008102-10004703/99-56                -S-00000000-99999999-0001"
    li = li + 1
    mTabSusep(li) = "8100000008120-********************          -S-00000000-99999999-0001"
    li = li + 1
    mTabSusep(li) = "8110000000000-********************          -S-00000000-99999999-0001"
    li = li + 1
    mTabSusep(li) = "8120000000000-********************          -S-00000000-99999999-0001"
    li = li + 1
    mTabSusep(li) = "8200000000000-********************          -S-00000000-99999999-0001"
    li = li + 1
    mTabSusep(li) = "8300000000000-********************          -S-00000000-99999999-0001"
    li = li + 1
    mTabSusep(li) = "9300000000000-15414.003269/97-13            -S-00000000-99999999-0001"
    li = li + 1
    mTabSusep(li) = "9310000000000-15414.003265/97-62            -S-00000000-99999999-0001"
    li = li + 1
    mTabSusep(li) = "9320000000000-15414.003267/97-98            -S-00000000-99999999-0001"
End Sub
