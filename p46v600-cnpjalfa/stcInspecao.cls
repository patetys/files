VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcInspecao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Num_Ind          As Double
Public Num_Item         As Double
Public Num_Insp         As Double
Public Num_CEP          As Long
Public tip_logr         As String
Public Nom_Logr         As String
Public Num_Logr         As Long
Public dsc_compl        As String
Public dsc_cid          As String
Public dsc_bairro       As String
Public uf               As String
Public cod_status       As Integer
Public dat_derie        As Long
Public cod_user_cad     As String
Public Cod_Clasf        As String
Public Cod_Compl_Clasf  As String
