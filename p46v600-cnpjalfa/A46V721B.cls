VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46V721B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit

' ======================================================================================================================================================

' Class Name: stcA46V721B --> Tabela Tab_Ped_Espec2
' Data/Hora: 15/09/2006 - 09:52:48
' Sistema: Multi
' Simony Sales
' ======================================================================================================================================================


' ======================================================================================================================================================
Public nosso_numero As String       'Nosso número.
Public Num_Ped As Long    'Número do pedido.
Public Cod_Org As Integer    'Código do órgão.
Public Cod_Ramo As e_CodRamo    'Código do ramo.
Public Num_Apol As Double     'Sem uso - número da apólice.
Public Num_Endo As Long     'Sem uso - número do endosso
Public Dsc_LocalExposicao As String             'Local de Exposicao.
Public Dsc_Objetos As String      'Memorando - Objeto Segurado
Public Dsc_RCobertos As String        'Memorando - Riscos Cobertos
Public Dsc_Caracteristica As String             'Memorando - Característica
Public Dsc_Perimetro As String        'Memorando - Perímetro da Cobertura
Public Dsc_Especificacao As String            'Memorando - Especificação do Serviço
Public Dsc_Coberturas As String         'Memorando - Coberturas
Public Dsc_Franquia As String       'Memorando - Franquias
Public Nom_Contratante As String          'Nome do Contratante
Public Dsc_Clausulas As String        'Memorando - Cláusulas
Public Dsc_Condicao As String       'Memorando - Condições
Public Dsc_OutrasInf As String        'Memorando - Outras Informações
Public Dsc_Observacao As String         'Memorando - Observação
Public Dat_Obra As Long     'Prazo da Obra
Public Dat_Manut As Long      'Prazo da Manutenção
'Public NUM_PI As Long   'Número do PI.
Public NUM_PI As Double   'Número do PI.
Public Ind_Impressao As Integer       'Indica se a apólice foi impressa ou não

' ======================================================================================================================================================
Private ObjStcPai As stcA46V702B
' ======================================================================================================================================================

Public Property Get Class_ID() As String
    Class_ID = CStr(nosso_numero)
End Property


Public Property Get gCopiaEspecificacao() As stcA46V721B
    Dim lobjstcFator As stcA46V721B
    Set lobjstcFator = New stcA46V721B
    With lobjstcFator
        .nosso_numero = nosso_numero
        .Num_Ped = .Num_Ped
        .Cod_Org = Cod_Org
        .Cod_Ramo = Cod_Ramo
        .Num_Apol = Num_Apol
        .Num_Endo = Num_Endo
        .Dsc_Objetos = Dsc_Objetos
        .Dsc_LocalExposicao = Dsc_LocalExposicao
        .Dsc_RCobertos = Dsc_RCobertos
        .Dsc_Caracteristica = Dsc_Caracteristica
        .Dsc_Perimetro = Dsc_Perimetro
        .Dsc_Especificacao = Dsc_Especificacao
        .Dsc_Coberturas = Dsc_Coberturas
        .Dsc_Franquia = Dsc_Franquia
        .Nom_Contratante = Nom_Contratante
        .Dsc_Clausulas = Dsc_Clausulas
        .Dsc_Condicao = Dsc_Condicao
        .Dsc_OutrasInf = Dsc_OutrasInf
        .Dsc_Observacao = Dsc_Observacao
        .Dat_Obra = Dat_Obra
        .Dat_Manut = Dat_Manut
        .NUM_PI = NUM_PI
    End With
    Set gCopiaFator = lobjstcFator
End Property

Public Property Set ObjPai(pobjstc As stcA46V702B)
    Set ObjStcPai = pobjstc
End Property

Public Property Get Cod_Unicidade() As Long
    If Not ObjStcPai Is Nothing Then
        Cod_Unicidade = ObjStcPai.Cod_Unicidade
    End If
End Property
