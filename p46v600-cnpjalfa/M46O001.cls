VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsM46O001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ======================================================================================================================================================
' Class Name: clsM46O001
' Sistema: SyasMult
' Analista: Marcio Oliveira
' Objetivo: Substitui o antigo módulo bas Adaptado para ser utilizado como classe gerenciavel.
' ======================================================================================================================================================

Public PermiteAlterarTaxaFinal As Boolean
Public PermiteQQDescontoComercial As Boolean
Public PermiteDigitarSemNormas As Boolean
Public PermiteAlterarFranquia As Boolean
Public PermiteSemSiscota As Boolean
Public gAlterarFranquia As Boolean
Public gPermiteSemSiscota As Boolean

Public PerfilEspecial As Long    '1 - Especial
'2 - Não especial

Public gProComBasica As Double    'Produto: comissão básica.

'tbm_juros
'Public gProJurMinAnt01        As Variant            'Produto: limite de juros.
'Public gProJurMinAnt02        As Variant            'Produto: limite de juros.
'Public gProJurMinAnt03        As Variant            'Produto: limite de juros.
'Public gProJurMinAnt04        As Variant            'Produto: limite de juros.
'Public gProJurMinAnt05        As Variant            'Produto: limite de juros.
'Public gProJurMinAnt06        As Variant            'Produto: limite de juros.
'Public gProJurMinAnt07        As Variant            'Produto: limite de juros.
'Public gProJurMinAnt08        As Variant            'Produto: limite de juros.
'Public gProJurMinAnt09        As Variant            'Produto: limite de juros.
'Public gProJurMinAnt10        As Variant            'Produto: limite de juros.
'Public gProJurMinAnt11        As Variant            'Produto: limite de juros.
'Public gProJurMinAnt12        As Variant            'Produto: limite de juros.
'Public gProJurMinPos01        As Variant            'Produto: limite de juros.
'Public gProJurMinPos02        As Variant            'Produto: limite de juros.
'Public gProJurMinPos03        As Variant            'Produto: limite de juros.
'Public gProJurMinPos04        As Variant            'Produto: limite de juros.
'Public gProJurMinPos05        As Variant            'Produto: limite de juros.
'Public gProJurMinPos06        As Variant            'Produto: limite de juros.
'Public gProJurMinPos07        As Variant            'Produto: limite de juros.
'Public gProJurMinPos08        As Variant            'Produto: limite de juros.
'Public gProJurMinPos09        As Variant            'Produto: limite de juros.
'Public gProJurMinPos10        As Variant            'Produto: limite de juros.
'Public gProJurMinPos11        As Variant            'Produto: limite de juros.
'Public gProJurMinPos12        As Variant            'Produto: limite de juros.

Public gProJurMinPadrao As Single
Public gProLimComissao As Double      'Produto: limite da comissão.
Public gProLimMarComercial As Single          'Produto: limite da margem comercial.
Public gProNome As String    'Produto: nome.
Public gProTipProtocolo As Byte         'Produto: tipo do protocolo.
Public gProVigInicio As Double    'Produto: vigência - início.
Public gProVigTermino As Double  'Produto: vigência - término.

Public Property Get Class_Name() As String
    Class_Name = "M46O001"
End Property


Public Property Get gTarifLMI() As Long
    gTarifLMI = 982
End Property
