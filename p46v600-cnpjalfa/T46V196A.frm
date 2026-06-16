VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmT46V196A 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista de Bicicletas"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10530
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   10530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSFlexGridLib.MSFlexGrid flxBicicleta 
      Height          =   2430
      Left            =   30
      TabIndex        =   0
      ToolTipText     =   "Coberturas calculadas."
      Top             =   30
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   4286
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
   Begin Threed.SSCommand cmdNovo 
      Height          =   360
      Left            =   30
      TabIndex        =   1
      ToolTipText     =   "Novo bem."
      Top             =   2610
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   635
      _StockProps     =   78
      Caption         =   "&Novo"
      ForeColor       =   8388608
   End
   Begin Threed.SSCommand cmdExcluir 
      Height          =   360
      Left            =   2130
      TabIndex        =   3
      ToolTipText     =   "Excluir o bem."
      Top             =   2610
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   635
      _StockProps     =   78
      Caption         =   "&Excluir"
      ForeColor       =   8388608
   End
   Begin Threed.SSCommand cmdAlterar 
      Height          =   360
      Left            =   1080
      TabIndex        =   2
      ToolTipText     =   "Volta à proposta."
      Top             =   2610
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   635
      _StockProps     =   78
      Caption         =   "&Alterar"
      ForeColor       =   8388608
   End
   Begin Threed.SSCommand cmdSair 
      Height          =   360
      Left            =   9480
      TabIndex        =   4
      ToolTipText     =   "Excluir o bem."
      Top             =   2610
      Width           =   1005
      _Version        =   65536
      _ExtentX        =   1773
      _ExtentY        =   635
      _StockProps     =   78
      Caption         =   "&Sair"
      ForeColor       =   8388608
   End
End
Attribute VB_Name = "frmT46V196A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objStcA46V708B As stcA46V708B

Public Sub Carregar_Formulario(ByRef pStcA46V708B As stcA46V708B)
    Set objStcA46V708B = pStcA46V708B
    Carregar_Grid pStcA46V708B.ColecaoPedLocBicicleta
End Sub

Private Sub Carregar_Grid(ByRef pColBicicleta As Collection)
    Call Form_Load

    Dim lItem As stcA46V744B

    For Each lItem In pColBicicleta
        lItem.Cod_Seq = flxBicicleta.Rows
        flxBicicleta.Rows = flxBicicleta.Rows + 1
        flxBicicleta.TextMatrix(flxBicicleta.Rows - 1, 0) = "" & lItem.Cod_Seq
        flxBicicleta.TextMatrix(flxBicicleta.Rows - 1, 1) = "" & lItem.Descricao
        flxBicicleta.TextMatrix(flxBicicleta.Rows - 1, 2) = "" & Format(lItem.Valor, "standard")
    Next
    
End Sub

Private Sub cmdAlterar_Click()
    Dim lItem As stcA46V744B
    
    If flxBicicleta.Row > 0 Then
        For Each lItem In objStcA46V708B.ColecaoPedLocBicicleta
            If lItem.Cod_Seq = flxBicicleta.TextMatrix(flxBicicleta.Row, 0) Then
                Exit For
            End If
        Next
        
        If lItem Is Nothing Then
            MsgBox "Nenhum item selecionado!", vbInformation, Me.Caption
        
        Else
            If lItem.ObjPai Is Nothing Then
                Set lItem.ObjPai = objStcA46V708B
            End If
            
            frmT46V197A.Alterar lItem
            frmT46V197A.Show vbModal, Me
            Carregar_Grid objStcA46V708B.ColecaoPedLocBicicleta
        End If
    Else
        MsgBox "Nenhum item selecionado!", vbInformation, Me.Caption
    End If
End Sub

Private Sub cmdExcluir_Click()
    Dim lItem As stcA46V744B
    
    Call gpHabSalvar
    
    If flxBicicleta.Row > 0 Then
    
        For Each lItem In objStcA46V708B.ColecaoPedLocBicicleta
            If lItem.Cod_Seq = flxBicicleta.TextMatrix(flxBicicleta.Row, 0) Then
                Exit For
            End If
        Next
        
        If lItem Is Nothing Then
            MsgBox "Nenhum item selecionado!", vbInformation, Me.Caption
        
        Else
            objStcA46V708B.ColecaoPedLocBicicleta.Remove lItem.Cod_Seq
            Carregar_Grid objStcA46V708B.ColecaoPedLocBicicleta
        End If
    Else
        MsgBox "Nenhum item selecionado!", vbInformation, Me.Caption
    End If
End Sub

Private Sub cmdNovo_Click()
    Dim lBicicleta As stcA46V744B
    Set lBicicleta = New stcA46V744B
    Set lBicicleta.ObjPai = objStcA46V708B
    
    frmT46V197A.Incluir lBicicleta
    frmT46V197A.Show vbModal, Me
    
    objStcA46V708B.ColecaoPedLocBicicleta.Add lBicicleta
    Carregar_Grid objStcA46V708B.ColecaoPedLocBicicleta
End Sub

Private Sub cmdSair_Click()
    Me.Hide
End Sub

Private Sub flxBicicleta_DblClick()
    Call cmdAlterar_Click
End Sub

Private Sub Form_Load()
    flxBicicleta.Rows = 1
    flxBicicleta.Cols = 3
    flxBicicleta.TextMatrix(0, 0) = "Item"
    flxBicicleta.ColWidth(0) = 580
    flxBicicleta.ColAlignment(0) = flexAlignRightCenter
    flxBicicleta.TextMatrix(0, 1) = "Descrição"
    flxBicicleta.ColWidth(1) = 7800
    flxBicicleta.ColAlignment(1) = flexAlignLeftCenter
    flxBicicleta.TextMatrix(0, 2) = "Valor"
    flxBicicleta.ColAlignment(2) = flexAlignRightCenter
    flxBicicleta.ColWidth(2) = 1500
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Set objStcA46V708B.ColecaoPedLocBicicleta = New Collection
End Sub
