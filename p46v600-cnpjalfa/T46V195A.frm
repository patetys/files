VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmT46V195A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Situação da verificação de pendência financeira"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6225
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSFlexGridLib.MSFlexGrid flxValidacao 
      Height          =   2580
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Coberturas calculadas."
      Top             =   0
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   4551
      _Version        =   393216
      Rows            =   5
      Cols            =   4
      FixedCols       =   0
      ForeColor       =   8388608
      ForeColorFixed  =   8388608
      ForeColorSel    =   12648447
      BackColorBkg    =   -2147483633
      GridColorFixed  =   4210752
      GridLines       =   2
      SelectionMode   =   1
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "frmT46V195A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function Carregar_Formulario(ByVal pNumero_Pi As String) As Boolean
    Carregar_Grid pNumero_Pi
    Me.Hide
    Carregar_Formulario = flxValidacao.Rows > 1
End Function

Private Sub Carregar_Grid(ByVal pNumero_Pi As String)
    Dim lSelect As String
    Dim lBDP00Multi As ADODB.Connection
    Dim lRstVerificacao As clsYasRecordSet
    Dim lMensagem As String
    
    lSelect = "SELECT  CASE TipoSistema WHEN 1 THEN 'SERASA' WHEN 2 THEN 'NORTRIX' WHEN 3 THEN 'SAP' WHEN 4 THEN 'BOA VISTA' ELSE 'OUTRO (' + CAST(TipoSistema AS VARCHAR) + ')' END TipoSistema,"
    lSelect = lSelect + "CASE LEN(Chave1) WHEN 11 THEN LEFT(Chave1, 3) + '.' + SUBSTRING(Chave1, 4,3) + '.' + SUBSTRING(Chave1, 7,3) + '-' + RIGHT(Chave1, 2) "
    lSelect = lSelect + "WHEN 13 THEN LEFT(Chave1, 1) + '.' + SUBSTRING(Chave1, 2,3) + '.' + SUBSTRING(Chave1, 5,3) + '/' + SUBSTRING(Chave1, 8,4) + '-' + RIGHT(Chave1, 2) "
    lSelect = lSelect + "WHEN 14 THEN LEFT(Chave1, 2) + '.' + SUBSTRING(Chave1, 3,3) + '.' + SUBSTRING(Chave1, 6,3) + '/' + SUBSTRING(Chave1, 9,4) + '-' + RIGHT(Chave1, 2) END Documento, "
    lSelect = lSelect + "CASE WHEN Dat_Processamento = 0 THEN 'NÃO PROCESSADO' ELSE RIGHT(Dat_Processamento, 2) + '/' + SUBSTRING(CAST(Dat_Processamento AS VARCHAR(8)), 5,2) +'/'+ LEFT(Dat_Processamento, 4) END Data_Proc "
    lSelect = lSelect + "FROM    p0044800..Tab_ConsDadosExternos "
    lSelect = lSelect + "WHERE cod_ramo in (0, 112, 113, 114) and num_pi <> 0 and num_pi = " & pNumero_Pi
    
    If gM46V999.gfAbrBasDados("P00MULTI", gM46V999.gTipP00MULTI, gM46V999.gEndP00MULTI, lBDP00Multi, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If
    
    If gM46V999.gfObtRegistro(lBDP00Multi, lSelect, lRstVerificacao, lMensagem) = False Then
        Call gM46V999.gpGraLog(2, lMensagem)
        Exit Sub
    End If
    
    flxValidacao.Rows = 1
    flxValidacao.Cols = 3
    flxValidacao.TextMatrix(0, 0) = "Tipo de Sistema"
    flxValidacao.ColWidth(0) = 1500
    flxValidacao.TextMatrix(0, 1) = "Documento"
    flxValidacao.ColWidth(1) = 1500
    flxValidacao.TextMatrix(0, 2) = "Data de Processamento"
    flxValidacao.ColWidth(2) = 2000
    
    Do While Not lRstVerificacao.EOF
        flxValidacao.Rows = flxValidacao.Rows + 1
        flxValidacao.TextMatrix(flxValidacao.Rows - 1, 0) = "" & lRstVerificacao!TipoSistema
        flxValidacao.TextMatrix(flxValidacao.Rows - 1, 1) = "" & lRstVerificacao!documento
        flxValidacao.TextMatrix(flxValidacao.Rows - 1, 2) = "" & lRstVerificacao!Data_Proc
        lRstVerificacao.MoveNext
    Loop
    
    Call gM46V999.gpFecha2(lBDP00Multi)
End Sub
