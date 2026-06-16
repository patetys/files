VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stcTabConsDadosExternos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'========================================
'Estrutura da tabela [P0044800].[dbo].[Tab_ConsDadosExternos]
'========================================
'Public NUM_PI           As Long            '[decimal](10, 0) NOT NULL ,   ***PK
Public NUM_PI           As Double            '[decimal](10, 0) NOT NULL ,   ***PK
Public Cod_Ramo         As Integer          '[decimal](2, 0) NOT NULL ,   ***PK
Public TipoSistema      As Integer          '[decimal](2, 0) NOT NULL ,   ***PK
Public Tip_Chave1       As Integer          '[varchar](1, 0) NOT NULL ,
Public Chave1           As String           '[varchar](50, 0) NOT NULL ,
Public Tip_Chave2       As Integer          '[varchar](1, 0) NOT NULL ,
Public Chave2           As String           '[varchar](50, 0) NOT NULL ,
Public Tip_Chave3       As Integer          '[varchar](1, 0) NOT NULL ,
Public Chave3           As String           '[varchar](50, 0) NOT NULL ,
Public Tip_Chave4       As Integer          '[varchar](1, 0) NOT NULL ,
Public Chave4           As String           '[varchar](50, 0) NOT NULL ,
Public Dat_Processamento  As Long           '[decimal] (8) NULL ,
Public Hor_Processamento  As Long           '[decimal] (8) NULL ,


Public Num_Item           As Integer        '[decimal](6, 0) NOT NULL ,   ***PK
Public Property Get Class_ID() As String
    Class_ID = CStr(NUM_PI) & _
               CStr(Num_Item) & _
               CStr(TipoSistema) & _
               CStr(Chave1)
End Property

