VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46v741B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Inspeção – Equipamentos
' ======================================================================================================================================================
' Class Name: stcA46v741B
' Data/Hora: 26/05/2015 - 11:00:00
' Sistema: P46V600
' Analista: Carlos Araujo
' ======================================================================================================================================================

Public nosso_numero         As String
Public NUM_PI               As Double     'PI.
Public Cod_Org              As Long
Public Cod_Ramo             As Long
Public Num_Ped              As Long
Public Num_Apol             As Double
Public Num_Endo             As Long       'Sem uso - número do endosso.
Public Tip_Agragado         As Long
Public Nom_Segurado         As String
Public Num_Cgc_Cpf          As String
Public Cod_Tip_Logr         As String     'Tipo do logradouro.
Public Nom_Logr             As String     'Logradouro.
Public Dsc_Compl_Logr       As String     'Complemento.
Public Nom_Bairro           As String     'Bairro.
Public Nom_Cid              As String     'Cidade.
Public Sig_UF               As String     'UF.
Public Num_CEP              As Long       'CEP.
Public Num_Fone             As Double     'Telefone.
Public Num_Logr             As Long       'Número do logradouro.
Public Tipo_Fone            As Long       'Tipo de telefone.
Public Num_DDD_Fone         As Integer
Public Num_Ramal            As Long
Public Tip_Pessoa           As Integer
Public Cod_Ramo_Ativ        As Integer    'Código do ramo de atividade.
Public Cod_Doc              As Integer    'Código do tipo de documento.
Public Num_Doc              As String     'Número do documento.
Public Org_Exp              As String     'Órgão expedidor.
Public Dat_Emis_Doc         As Long       'Data da emissão do documento (yyyymmdd).
Public DAT_Nasc             As Long
Public Nom_Empresa          As String     'Nome da empresa onde trabalha.
Public Sexo                 As String     'Sexo do segurado: M -masculino , F - Feminino.
Public Renda_Fatura         As Double

Public DAT_GRAV                 As Long                     'DAT GRAV
Public Cod_User                 As String                   'COD USER
Public DIG_BLOQUETO             As Byte
Private mDigito As String
Private mBloquetoDigito As Long

Public ColecaoPedAgregados As Collection

' ======================================================================================================================================================
Private ObjStcPai As stcA46V702B
' ======================================================================================================================================================
Public Property Get Class_ID() As String
    Class_ID = CStr(nosso_numero)
End Property

Public Property Set ObjPai(pobjstc As stcA46V702B)
    Set ObjStcPai = pobjstc
End Property

Public Property Get Cod_Unicidade() As Long
    If Not ObjStcPai Is Nothing Then
        Cod_Unicidade = ObjStcPai.Cod_Unicidade
    End If
End Property


