Attribute VB_Name = "basM46O001"
Option Explicit

Public PermiteAlterarTaxaFinal As Boolean
Public PermiteQQDescontoComercial As Boolean
Public PermiteDigitarSemNormas As Boolean
Public PermiteAlterarFranquia As Boolean
Public PermiteSemSiscota      As Boolean
Public gAlterarFranquia      As Boolean
Public gPermiteSemSiscota    As Boolean

Public gPerfilEspecial  As Long         '1 - Especial
                                        '2 - Não especial
Public Const gTarifLMI = 982


Public gProComBasica        As Double           'Produto: comissão básica.

Public gProJurMinAnt01      As Single           'Produto: limite de juros.
Public gProJurMinAnt02      As Single           'Produto: limite de juros.
Public gProJurMinAnt03      As Single           'Produto: limite de juros.
Public gProJurMinAnt04      As Single           'Produto: limite de juros.
Public gProJurMinAnt05      As Single           'Produto: limite de juros.
Public gProJurMinAnt06      As Single           'Produto: limite de juros.
Public gProJurMinAnt07      As Single           'Produto: limite de juros.
Public gProJurMinAnt08      As Single           'Produto: limite de juros.
Public gProJurMinAnt09      As Single           'Produto: limite de juros.
Public gProJurMinAnt10      As Single           'Produto: limite de juros.
Public gProJurMinAnt11      As Single           'Produto: limite de juros.
Public gProJurMinAnt12      As Single           'Produto: limite de juros.
Public gProJurMinPos01      As Single           'Produto: limite de juros.
Public gProJurMinPos02      As Single           'Produto: limite de juros.
Public gProJurMinPos03      As Single           'Produto: limite de juros.
Public gProJurMinPos04      As Single           'Produto: limite de juros.
Public gProJurMinPos05      As Single           'Produto: limite de juros.
Public gProJurMinPos06      As Single           'Produto: limite de juros.
Public gProJurMinPos07      As Single           'Produto: limite de juros.
Public gProJurMinPos08      As Single           'Produto: limite de juros.
Public gProJurMinPos09      As Single           'Produto: limite de juros.
Public gProJurMinPos10      As Single           'Produto: limite de juros.
Public gProJurMinPos11      As Single           'Produto: limite de juros.
Public gProJurMinPos12      As Single           'Produto: limite de juros.
Public gProJurMinPadrao     As Single
Public gProLimComissao      As Double           'Produto: limite da comissão.
Public gProLimMarComercial  As Single           'Produto: limite da margem comercial.
Public gProNome             As String           'Produto: nome.
Public gProPreMinEndosso    As Double           'Produto: prêmio mínimo por parcela de endosso.
Public gProTipProtocolo     As Byte             'Produto: tipo do protocolo.
Public gProVigInicio       As Double       'Produto: vigência - início.
Public gProVigTermino      As Double       'Produto: vigência - término.

