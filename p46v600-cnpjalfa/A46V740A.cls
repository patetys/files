VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA46v740A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Inspeção – Equipamentos
' ======================================================================================================================================================
' Class Name: clsA46v740A
' Data/Hora: 02/09/2014 - 14:00:00
' Sistema: P46V600
' Analista: Toshio Okamoto
' 21/09/2015
' Inclusão de Novos Campos
' ======================================================================================================================================================

Public ConexaoBD                As ADODB.Connection  ' Conexão com o Banco de Dados
Public Colecao                  As Collection        ' Coleção de objetos do tipo stcA46v730B
Public UltimoErro               As String            ' Descrição do último erro gerado

Private mClsM46V999 As clsM46V999

Public Property Get Class_Name() As String
    Class_Name = "A46v740A"
End Property

Private Function mfFormatarConsultaSQL(pobjstc As stcA46v740B) As String

Dim strTMP As String
    
    strTMP = "Select * from Tab_Trans_Insp_Equipamentos where Dat_Trans >= " & pobjstc.Dat_Trans
    strTMP = strTMP & " and Dat_Trans <= " & pobjstc.Dat_Trans
    
    mfFormatarConsultaSQL = strTMP
    
End Function

Private Function mfFormatarInclusaoSQL(pobjstc As stcA46v740B) As String

Dim strTMP As String
 
    strTMP = " INSERT INTO Tab_Trans_Insp_Equipamentos ( "
    strTMP = strTMP & " NOSSO_NUMERO,"
    strTMP = strTMP & " NUM_PED,"
    strTMP = strTMP & " NUM_ITEM,"
    strTMP = strTMP & " NUM_APOL_ANT,"
    strTMP = strTMP & " NOM_SEGURADO,"
    strTMP = strTMP & " NUM_TEL,"
    strTMP = strTMP & " NUM_CGC_CPF,"
    strTMP = strTMP & " COD_RAMO,"
    strTMP = strTMP & " COD_ORG,"
    strTMP = strTMP & " DAT_INI_VIG,"
    strTMP = strTMP & " DAT_TER_VIG,"
    strTMP = strTMP & " DAT_TRANS,"
    strTMP = strTMP & " COD_IDEN_PESSOA,"
    strTMP = strTMP & " COD_UNID_PROD,"
    strTMP = strTMP & " COD_PROD,"
    
    strTMP = strTMP & " [NUM_EQUIP],"
    strTMP = strTMP & " [NOME_EQUIP],"
    strTMP = strTMP & " [TIPO_EQUIP],"
    strTMP = strTMP & " [CLASSE_EQUIP],"
    strTMP = strTMP & " [ANO_FABRICACAO_EQUIP],"
    strTMP = strTMP & " [MARCA_EQUIP],"
    strTMP = strTMP & " [MODELO_EQUIP],"
    strTMP = strTMP & " [NUM_SERIE_EQUIP],"
    strTMP = strTMP & " [CHASSI_EQUIP],"
    strTMP = strTMP & " [NOME_CONTATO1],"
    strTMP = strTMP & " [TEL_CONTATO1],"
    strTMP = strTMP & " [EMAIL_CONTATO1],"
    strTMP = strTMP & " [NOME_CONTATO2],"
    strTMP = strTMP & " [TEL_CONTATO2],"
    strTMP = strTMP & " [EMAIL_CONTATO2],"
    strTMP = strTMP & " [NOME_CONTATO3],"
    strTMP = strTMP & " [TEL_CONTATO3],"
    strTMP = strTMP & " [EMAIL_CONTATO3],"
    strTMP = strTMP & " [OBSERVACOES],"
    strTMP = strTMP & " [LOGRADOURO],"
    strTMP = strTMP & " [LOGR_COMPL],"
    strTMP = strTMP & " [LOGR_NUMERO],"
    strTMP = strTMP & " [BAIRRO],"
    strTMP = strTMP & " [CIDADE],"
    strTMP = strTMP & " [UF],"
    strTMP = strTMP & " [CEP],"
    strTMP = strTMP & " [NUM_PI],"
    
    strTMP = strTMP & " NUM_INSP,"
    strTMP = strTMP & " COD_STATUS_I,"
    strTMP = strTMP & " DAT_CONFIRM,"
    strTMP = strTMP & " DAT_FINAL,"
    strTMP = strTMP & " DAT_CANCEL"
    strTMP = strTMP & ") VALUES ("
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Num_Ped) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Num_Apol_Ant) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Nom_Segurado) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Num_Tel) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Num_Cgc_Cpf) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Ramo) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Org) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Ini_Vig) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Ter_Vig) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Trans) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Iden_Pessoa) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Unid_Prod) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Prod) & ","
    
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.NUM_EQUIP) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.NOME_EQUIP) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.TIPO_EQUIP) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.CLASSE_EQUIP) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.ANO_FABRICACAO_EQUIP) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.MARCA_EQUIP) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.MODELO_EQUIP) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.NUM_SERIE_EQUIP) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.CHASSI_EQUIP) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.NOME_CONTATO1) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.TEL_CONTATO1) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.EMAIL1) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.NOME_CONTATO2) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.TEL_CONTATO2) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.EMAIL2) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.NOME_CONTATO3) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.TEL_CONTATO3) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.EMAIL3) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Observacoes) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.LOGRADOURO) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.LOGR_COMPL) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.LOGR_NUMERO) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.BAIRRO) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.CIDADE) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.UF) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.CEP) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.NUM_PI) & ","
    
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Num_Insp) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Cod_Status_I) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Confirm) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Final) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Dat_Cancel) & ")"

    mfFormatarInclusaoSQL = strTMP

End Function

Private Function mfFormatarInclusaoSQL_COMPL(pobjstc As stcA46v740B) As String

Dim strTMP As String
 
    strTMP = " INSERT INTO Inspecao..Tab_Insp_Equip_Compl ( "
    strTMP = strTMP & " NOSSO_NUMERO,"
    strTMP = strTMP & " NUM_PED,"
    strTMP = strTMP & " NUM_ITEM,"
    strTMP = strTMP & " NUM_PI,"
    strTMP = strTMP & " NUM_INSP,"
    
    strTMP = strTMP & " [NOME_CONTATO1],"
    strTMP = strTMP & " [TEL_CONTATO1],"
    strTMP = strTMP & " [EMAIL_CONTATO1],"
    strTMP = strTMP & " [NOME_CONTATO2],"
    strTMP = strTMP & " [TEL_CONTATO2],"
    strTMP = strTMP & " [EMAIL_CONTATO2],"
    strTMP = strTMP & " [NOME_CONTATO3],"
    strTMP = strTMP & " [TEL_CONTATO3],"
    strTMP = strTMP & " [EMAIL_CONTATO3],"
    strTMP = strTMP & " [OBSERVACOES]"
    
    strTMP = strTMP & ") VALUES ("
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Num_Ped) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.NUM_PI) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Num_Insp) & ","
    
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.NOME_CONTATO1) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.TEL_CONTATO1) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.EMAIL1) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.NOME_CONTATO2) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.TEL_CONTATO2) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.EMAIL2) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.NOME_CONTATO3) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.TEL_CONTATO3) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.EMAIL3) & ","
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Observacoes) & ")"

    mfFormatarInclusaoSQL_COMPL = strTMP

End Function

Private Function mfFormatarExclusaoSQL(pobjstc As stcA46v740B) As String

Dim strTMP As String
 
    strTMP = " DELETE from Tab_Trans_Insp_Equipamentos "
    strTMP = strTMP & " where Nosso_Numero = "
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.nosso_numero)
    strTMP = strTMP & " and Num_Ped = "
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Num_Ped)
    strTMP = strTMP & " and Num_Item = "
    strTMP = strTMP & M46V999.gfFormatarCamposSQL(pobjstc.Num_Item)
    strTMP = strTMP & " AND Cod_Status_I = 1"

    mfFormatarExclusaoSQL = strTMP
    
End Function

Private Sub mpLimparColecao()
    Set Colecao = Nothing
    Set Colecao = New Collection
End Sub
Public Function gfIncluir(pobjstc As stcA46v740B) As Boolean
    
    Dim lConexaoSQL As ADODB.Connection
    Dim sSql        As String
    Dim pMensagem   As String
    Dim lRegAlt     As Long
 
    On Local Error GoTo ControleErros
    
    sSql = mfFormatarExclusaoSQL(pobjstc)
    
    If M46V999.gfAbrBasDados("RAMOSDIVERSOS", M46V999.gTipDIVERSOS, M46V999.gEndDIVERSOS, lConexaoSQL, pMensagem) = False Then
        Exit Function
    End If
    
    If M46V999.gfExeSQL(lConexaoSQL, sSql, UltimoErro, lRegAlt) <> 0 Then GoTo ControleErros
    
   
    sSql = mfFormatarInclusaoSQL(pobjstc)
       
    If M46V999.gfExeSQL(lConexaoSQL, sSql, UltimoErro) <> 0 Then GoTo ControleErros
    
    sSql = mfFormatarInclusaoSQL_COMPL(pobjstc)
    
    If M46V999.gfExeSQL(lConexaoSQL, sSql, UltimoErro) <> 0 Then GoTo ControleErros
    
    Colecao.Add Item:=pobjstc, Key:=pobjstc.Class_ID
    
    gfIncluir = True
    
    Exit Function
    
ControleErros:
    gfIncluir = False
 '   If UCase(UltimoErro) Like "*" & UCase("duplicate key") & "*" Then
 '       UltimoErro = "DUPLICADO"
 '       gfIncluir = True
 '   Else
        If Not M46V999.gfPreenchido(UltimoErro) Then
             UltimoErro = Err.Description
        End If
 '   End If
    
End Function

Private Sub Class_Initialize()

    Set Colecao = New Collection
    Set mClsM46V999 = Nothing
End Sub
Private Sub Class_Terminate()
    Set Colecao = Nothing
    Set mClsM46V999 = Nothing
End Sub

Public Property Set M46V999(pClsM46V999 As clsM46V999)
    Set mClsM46V999 = pClsM46V999
End Property

Public Property Get M46V999() As clsM46V999
    If mClsM46V999 Is Nothing Then
         Err.Raise 91, Me.Class_Name, "Objeto M46V999 não atribuído para a classe (" & Me.Class_Name & ")"
         '***Esta Classe DEVE utilizar uma instancia da classe clsM46V999 já
         '   iniciada no escopo onde esta classe foi instanciada.
         'Atribua a o objeto (clsM46V999 ja instanciado) no momento que eu inicializou esta classe (atraves do Property SET)!!!
    End If
    
Retorno:
    Set M46V999 = mClsM46V999
End Property

