Attribute VB_Name = "xxxS00U008"
Option Explicit
'PRM-COD-USER...gNomUsuario.
Public gS00U008_CodClaOcupacao  As String       'PRM-COD-CLASSE-OCUP...Código da classe de ocupação.
Public gS00U008_CodCobertura    As String       'PRM-COD-COBERT........Código da cobertura.
Public gS00U008_CodConstrucao   As String       'PRM-COD-CONSTR........Código da construção.
Public gS00U008_CodCorretor(3)  As String       'PRM-COD-CORR..........Código do corretor (até 3).
Public gS00U008_CodEstipulante  As String       'PRM-COD-ESTIP.........Código do estipulante.
Public gS00U008_CodSegurado     As String       'PRM-COD-SEGURADO......Código do segurado.
Public gS00U008_Form            As String       'PRM-NOM-TELA..........Nome do form.
Public gS00U008_Mensagem        As String       'PRM-DSC-ERRO..........Mensagem de erro.
Public gS00U008_NomCampo        As String       'PRM-NOM-CAMPO.........Nome do campo do form.
Public gS00U008_Opcao           As String       'PRM-COD-TIP-OPC.......1 = carrega tabela
                                                '                      2 = consistência.
Public gS00U008_Ramo            As String       'PRM-COD-RAMO..........Código do ramo.
Public gS00U008_Retorno         As Byte         'PRM-COD-SIT...........Código de retorno.
                                                '                      0  = OK.
                                                '                      88 = não encontrou limite.
                                                '                      99 = erro de limite.
Public gS00U008_TabLimite       As String       'PRM-TAB-LIM...........Tabela de limites.
Public gS00U008_ValCampo        As Double       'PRM-VAL-CAMPO.........Valor do campo.

Private mAntRamo                As String       'PRM-COD-RAMO-BAK.
Private mAuxTabLimite           As Integer      'TAB-LIM-HOST.
Private mAuxValLimMinimo        As Double       'VAL-LIM-MIN-ERRO-WS.
Private mAuxValLimMaximo        As Double       'VAL-LIM-MAX-ERRO-WS.
Private mLimCodClasse()         As String       'TAB-LIM: COD-CLASSE-WS.
Private mLimCodCobertura()      As String       'TAB-LIM: COD-COBERT-WS.
Private mLimCodConstrucao()     As String       'TAB-LIM: COD-CONSTR-WS.
Private mLimCodCorretor()       As String       'TAB-LIM: COD-CORR-WS.
Private mLimCodSegurado()       As String       'TAB-LIM: COD-SEGURADO-WS.
Private mLimCodUsuario()        As String       'TAB-LIM: COD-USER-TAB-WS.
Private mLimMaximo()            As Double       'TAB-LIM: VAL-LIM-MAX-WS.
Private mLimMinimo()            As Double       'TAB-LIM: VAL-LIM-MIN-WS.
Private mLimNomCampo()          As String       'TAB-LIM: NOM-CAMPO-WS.
Private mLimNomForm()           As String       'TAB-LIM: NOM-TELA-WS.
Private mLimRamo()              As String       'TAB-LIM: COD-RAMO-WS.
Public Sub gpS00U008()
    '1. Início (001-00-INICIALIZA).
    If IsNumeric(gS00U008_TabLimite) Then
        If CInt(gS00U008_TabLimite) = 0 Then
            mAuxTabLimite = 1
        Else
            mAuxTabLimite = gS00U008_TabLimite
        End If
    Else
        mAuxTabLimite = 1
    End If
    If gfPreenchido(gS00U008_Ramo) Then
        mAntRamo = gS00U008_Ramo
    Else
        mAntRamo = ""
    End If
    mAuxValLimMinimo = 0
    mAuxValLimMaximo = 0

    '2. Consiste parâmetros (003-00-CONSISTE-PARM).
    If gS00U008_Opcao <> "1" And gS00U008_Opcao <> "2" Then
        gS00U008_Retorno = 99
        gS00U008_Mensagem = "Opção diferente de 1 e 2."
        If gfPreenchido(mAntRamo) Then
            gS00U008_Ramo = mAntRamo
        Else
            gS00U008_Ramo = ""
        End If
        Exit Sub
    End If
    If Not gfPreenchido(gS00U008_Form) Then
        gS00U008_Retorno = 99
        gS00U008_Mensagem = "Nome do form não informado."
        If gfPreenchido(mAntRamo) Then
            gS00U008_Ramo = mAntRamo
        Else
            gS00U008_Ramo = ""
        End If
        Exit Sub
    End If
    If gS00U008_Opcao = "2" And Not gfPreenchido(gS00U008_NomCampo) Then
        gS00U008_Retorno = 99
        gS00U008_Mensagem = "Nome do campo não informado."
        If gfPreenchido(mAntRamo) Then
            gS00U008_Ramo = mAntRamo
        Else
            gS00U008_Ramo = ""
        End If
        Exit Sub
    End If

    '3. Processa (005-00-PROCESSA).
    Do  '005-01-PROCESSA.
        gS00U008_Mensagem = ""
        'a) Se opção 1, carrega tabela.
        If gS00U008_Opcao = "1" Then
            Call mpCarTabela
            If gfPreenchido(mAntRamo) Then
                gS00U008_Ramo = mAntRamo
            Else
                gS00U008_Ramo = ""
            End If
            Exit Sub
        End If
        'b) Verifica dados.
        Call mpVerDados
        If gS00U008_Retorno = 99 Then
            gS00U008_Mensagem = "Valor fora do limite (de " & mAuxValLimMinimo & " a " & _
                                mAuxValLimMaximo & ")."
            Exit Do
        Else
            'IF COD-CLASSE-OCUP NOT = 0
            '   MOVE 0 TO COD-CLASSE-OCUP
            '   GO TO 005-01-PROCESSA
            'ELSE
                Exit Do
            'END-IF
        End If
    Loop    '(005-99-PROCESSA).
    If gfPreenchido(mAntRamo) Then
        gS00U008_Ramo = mAntRamo
    Else
        gS00U008_Ramo = ""
    End If
End Sub
Private Sub mpCarTabela()
    'Procedure: carrega tabela (007-00-CARREGA-TABELA).

    Dim lI      As Integer      'Índice.
    Dim rsLIM   As Recordset

    '1. Obtém dados de limite.
    gSelect = "SELECT * FROM YAS.TAB_LIM WHERE NOM_TELA = " & gS00U008_Form & " AND TAB_LIM = " & _
              mAuxTabLimite
    Set rsLIM = gfObtRegistro(cnn, gTipXXXXXX, gSelect)
    If rsLIM.RecordCount = 0 Then
        gS00U008_Retorno = 99
        gS00U008_Mensagem = "Erro no carregamento da tabela."
        rsLIM.Close
        Set rsLIM = Nothing
        Exit Sub
    End If
    
    '2. Carrega dados na tabela (007-03-CARREGA-TABELA).
    lI = 0
    While Not rsLIM.EOF
        'Comandos inibidos no programa Cobol.
        'If gfPreenchido(gS00U008_Ramo) Then
        '    If IsNumeric(gS00U008_Ramo) Then
        '        If CInt(gS00U008_Ramo) <> 0 Then
        '            If gS00U008_Ramo <> rsLIM!COD_RAMO Then
        '                GO TO 007-05-CARREGA-TABELA (processa próximo registro).
        '            End If
        '        End If
        '    End If
        'End If
        lI = lI + 1
        ReDim Preserve mLimNomForm(lI)
        ReDim Preserve mLimCodCorretor(lI)
        ReDim Preserve mLimCodSegurado(lI)
        ReDim Preserve mLimCodUsuario(lI)
        ReDim Preserve mLimRamo(lI)
        ReDim Preserve mLimCodCobertura(lI)
        ReDim Preserve mLimCodClasse(lI)
        ReDim Preserve mLimCodConstrucao(lI)
        ReDim Preserve mLimNomCampo(lI)
        ReDim Preserve mLimMinimo(lI)
        ReDim Preserve mLimMaximo(lI)
        mLimNomForm(lI) = rsLIM!NOM_TELA
        mLimCodCorretor(lI) = rsLIM!Cod_Corr
        mLimCodSegurado(lI) = rsLIM!COD_SEGURADO
        mLimCodUsuario(lI) = rsLIM!Cod_User
        mLimRamo(lI) = rsLIM!COD_RAMO
        mLimCodCobertura(lI) = rsLIM!Cod_Cobert
        mLimCodClasse(lI) = rsLIM!COD_CLASSE_OCUP
        mLimCodConstrucao(lI) = rsLIM!COD_CONSTR
        mLimNomCampo(lI) = rsLIM!NOM_CAMPO
        mLimMinimo(lI) = rsLIM!VAL_LIM_MIN
        mLimMaximo(lI) = rsLIM!VAL_LIM_MAX
        rsLIM.MoveNext
    Wend
    rsLIM.Close
    Set rsLIM = Nothing
    gS00U008_Retorno = 0
End Sub
Private Sub mpVerCorretor()
    'Procedure: verifica corretor (011-00-CONSISTE-CORRETOR).

    Dim lI1 As Byte     'Utilizado com For...Next.
    Dim li2 As Integer  'Utilizado com For...Next (para tabela de limites).

    For lI1 = 1 To 3
        If gS00U008_CodCorretor(lI1) = 0 Then
            Exit For
        End If
        For li2 = 1 To UBound(mLimCodCorretor)
            If mLimCodCorretor(li2) = gS00U008_CodCorretor(lI1) And _
               mLimCodSegurado(li2) = 0 And _
               Not gfPreenchido(mLimCodUsuario(li2)) And _
               mLimRamo(li2) = gS00U008_Ramo And _
               mLimCodCobertura(li2) = gS00U008_CodCobertura And _
               mLimNomCampo(li2) = gS00U008_NomCampo And _
               mLimCodClasse(li2) = gS00U008_CodClaOcupacao And _
               mLimCodConstrucao(li2) = gS00U008_CodConstrucao Then
                If gS00U008_ValCampo < mLimMinimo(li2) Or gS00U008_ValCampo > mLimMaximo(li2) Then
                    gS00U008_Retorno = 99
                    mAuxValLimMinimo = mLimMinimo(li2)
                    mAuxValLimMaximo = mLimMaximo(li2)
                Else
                    gS00U008_Retorno = 0
                End If
                Exit Sub
            End If
        Next li2
    Next lI1
    gS00U008_Retorno = 88
End Sub
Private Sub mpVerDados()
    'Procedure: verifica dados.

    '1. Consiste usuário.
    Call mpVerUsuario
    If gS00U008_Retorno = 0 Or gS00U008_Retorno = 99 Then
        Exit Sub
    End If

    '2. Consiste segurado.
    Call mpVerSegurado
    If gS00U008_Retorno = 0 Or gS00U008_Retorno = 99 Then
        Exit Sub
    End If

    '3. Consiste corretor.
    Call mpVerCorretor
    If gS00U008_Retorno = 0 Or gS00U008_Retorno = 99 Then
        Exit Sub
    End If

    '4. Consiste estipulante.
    Call mpVerEstipulante
    If gS00U008_Retorno = 0 Or gS00U008_Retorno = 99 Then
        Exit Sub
    End If

    '5. Consiste importância segurada básica.
    Call mpVerIS_Basica
    If gS00U008_Retorno = 0 Then
        Exit Sub
    End If

    '6. Se não encontrou limite, verifica novamente com ramo igual a 000.
    If gS00U008_Retorno = 88 Then
        gS00U008_Mensagem = ""
        gS00U008_Ramo = ""
        Call mpVerUsuario
        If gS00U008_Retorno = 0 Or gS00U008_Retorno = 99 Then
            Exit Sub
        End If
        Call mpVerSegurado
        If gS00U008_Retorno = 0 Or gS00U008_Retorno = 99 Then
            Exit Sub
        End If
        Call mpVerCorretor
        If gS00U008_Retorno = 0 Or gS00U008_Retorno = 99 Then
            Exit Sub
        End If
        Call mpVerEstipulante
        If gS00U008_Retorno = 0 Or gS00U008_Retorno = 99 Then
            Exit Sub
        End If
        Call mpVerIS_Basica
        If gS00U008_Retorno = 0 Then
            Exit Sub
        End If
    End If
End Sub
Private Sub mpVerEstipulante()
    'Procedure: verifica estipulante (013-00-CONSISTE-ESTIPUL).

    Dim lI  As Integer      'Utilizado com For...Next.

    If Not gfPreenchido(gS00U008_CodEstipulante) Then
        Exit Sub
    End If
    For lI = 1 To UBound(mLimCodCorretor)
        If mLimCodCorretor(lI) = 0 And _
           mLimCodSegurado(lI) = gS00U008_CodEstipulante And _
           Not gfPreenchido(mLimCodUsuario(lI)) And _
           mLimRamo(lI) = gS00U008_Ramo And _
           mLimCodCobertura(lI) = gS00U008_CodCobertura And _
           mLimNomCampo(lI) = gS00U008_NomCampo And _
           mLimCodClasse(lI) = gS00U008_CodClaOcupacao And _
           mLimCodConstrucao(lI) = gS00U008_CodConstrucao Then
            If gS00U008_ValCampo < mLimMinimo(lI) Or gS00U008_ValCampo > mLimMaximo(lI) Then
                gS00U008_Retorno = 99
                mAuxValLimMinimo = mLimMinimo(lI)
                mAuxValLimMaximo = mLimMaximo(lI)
            Else
                gS00U008_Retorno = 0
            End If
            Exit Sub
        End If
    Next lI
    gS00U008_Retorno = 88
End Sub
Private Sub mpVerIS_Basica()
    'Procedure: verifica importância segurada básica (009-00-CONSISTE-BASICA).

    Dim lI  As Integer      'Utilizado com For...Next.

    For lI = 1 To UBound(mLimCodCorretor)
        If mLimCodCorretor(lI) = 0 And _
           mLimCodSegurado(lI) = 0 And _
           Not gfPreenchido(mLimCodUsuario(lI)) And _
           mLimRamo(lI) = gS00U008_Ramo And _
           mLimCodCobertura(lI) = gS00U008_CodCobertura And _
           mLimNomCampo(lI) = gS00U008_NomCampo And _
           mLimCodClasse(lI) = gS00U008_CodClaOcupacao And _
           mLimCodConstrucao(lI) = gS00U008_CodConstrucao Then
            If gS00U008_ValCampo < mLimMinimo(lI) Or gS00U008_ValCampo > mLimMaximo(lI) Then
                gS00U008_Retorno = 99
                mAuxValLimMinimo = mLimMinimo(lI)
                mAuxValLimMaximo = mLimMaximo(lI)
            Else
                gS00U008_Retorno = 0
            End If
            Exit Sub
        End If
    Next lI
    gS00U008_Retorno = 88
    gS00U008_Mensagem = "Limite não encontrado. Verificar com Desis."
End Sub
Private Sub mpVerSegurado()
    'Procedure: verifica segurado (015-00-CONSISTE-SEGURADO).

    Dim lI  As Integer      'Utilizado com For...Next.

    If Not gfPreenchido(gS00U008_CodSegurado) Then
        Exit Sub
    End If
    For lI = 1 To UBound(mLimCodCorretor)
        If mLimCodCorretor(lI) = 0 And _
           mLimCodSegurado(lI) = gS00U008_CodSegurado And _
           Not gfPreenchido(mLimCodUsuario(lI)) And _
           mLimRamo(lI) = gS00U008_Ramo And _
           mLimCodCobertura(lI) = gS00U008_CodCobertura And _
           mLimNomCampo(lI) = gS00U008_NomCampo And _
           mLimCodClasse(lI) = gS00U008_CodClaOcupacao And _
           mLimCodConstrucao(lI) = gS00U008_CodConstrucao Then
            If gS00U008_ValCampo < mLimMinimo(lI) Or gS00U008_ValCampo > mLimMaximo(lI) Then
                gS00U008_Retorno = 99
                mAuxValLimMinimo = mLimMinimo(lI)
                mAuxValLimMaximo = mLimMaximo(lI)
            Else
                gS00U008_Retorno = 0
            End If
        End If
    Next lI
    gS00U008_Retorno = 88
End Sub
Private Sub mpVerUsuario()
    'Procedure: verifica usuário (017-00-CONSISTE-USUARIO).

    Dim lI  As Integer      'Utilizado com For...Next.


    If Not gfPreenchido(gNomUsuario) Then
        Exit Sub
    End If
    For lI = 1 To UBound(mLimCodCorretor)
        If mLimCodCorretor(lI) = 0 And _
           mLimCodSegurado(lI) = 0 And _
           mLimCodUsuario(lI) = gNomUsuario And _
           mLimRamo(lI) = gS00U008_Ramo And _
           mLimCodCobertura(lI) = gS00U008_CodCobertura And _
           mLimNomCampo(lI) = gS00U008_NomCampo And _
           mLimCodClasse(lI) = gS00U008_CodClaOcupacao And _
           mLimCodConstrucao(lI) = gS00U008_CodConstrucao Then
            If gS00U008_ValCampo < mLimMinimo(lI) Or gS00U008_ValCampo > mLimMaximo(lI) Then
                gS00U008_Retorno = 99
                mAuxValLimMinimo = mLimMinimo(lI)
                mAuxValLimMaximo = mLimMaximo(lI)
            Else
                gS00U008_Retorno = 0
            End If
        End If
    Next lI
    gS00U008_Retorno = 88
End Sub
