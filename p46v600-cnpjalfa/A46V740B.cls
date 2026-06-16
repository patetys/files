VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcA46v740B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Inspeção – Equipamentos
' ======================================================================================================================================================
' Class Name: stcA46v740B
' Data/Hora: 02/09/2014 - 14:00:00
' Sistema: P46V600
' Analista: Toshio Okamoto
' 21/09/2015
' Inclusão de Novos Campos
' ======================================================================================================================================================

Public nosso_numero         As String
Public Num_Ped              As Long
Public Num_Item             As Long
Public Num_Apol_Ant         As Double
Public Nom_Segurado         As String
Public Num_Tel              As String
Public Num_Cgc_Cpf          As String
Public Cod_Ramo             As Long
Public Cod_Org              As Long
Public Dat_Ini_Vig          As Double
Public Dat_Ter_Vig          As Double
Public Dat_Trans            As Double
Public Cod_Iden_Pessoa      As String
Public Cod_Unid_Prod        As Long
Public Cod_Prod             As Long

Public NUM_EQUIP            As Long
Public NOME_EQUIP           As String
Public TIPO_EQUIP           As Long
Public CLASSE_EQUIP         As Long
Public ANO_FABRICACAO_EQUIP As Long
Public MARCA_EQUIP          As String
Public MODELO_EQUIP         As String
Public NUM_SERIE_EQUIP      As String
Public CHASSI_EQUIP         As String
Public NOME_CONTATO1        As String
Public TEL_CONTATO1         As Double
Public EMAIL1               As String
Public NOME_CONTATO2        As String
Public TEL_CONTATO2         As Double
Public EMAIL2               As String
Public NOME_CONTATO3        As String
Public TEL_CONTATO3         As Double
Public EMAIL3               As String
Public Observacoes          As String
Public LOGRADOURO           As String
Public LOGR_COMPL           As String
Public LOGR_NUMERO          As Long
Public BAIRRO               As String
Public CIDADE               As String
Public UF                   As String
Public CEP                  As Double
Public NUM_PI               As Double

Public Num_Insp             As Long
Public Cod_Status_I         As Integer
Public Dat_Confirm          As Double
Public Dat_Final            As Double
Public Dat_Cancel           As Double

Public Property Get Class_ID() As String
    Class_ID = CStr(nosso_numero)
End Property

