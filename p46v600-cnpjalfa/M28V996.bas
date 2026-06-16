Attribute VB_Name = "basM28V996"
Option Explicit
'-------------------------------------------------------------------------------------------
'Autor: Márcio Oliveira
'Data:  01/2006
'Objetivo: Módulo utilizado para:
'                - Manipulação do Registro do Windows
'                - Retornar o caminho completo para os programas do Office (utilizados para automatizar a abertura de documentos do office)
'
'   Function gfRegistry_Read    Retorna um valor do registro, ref. a chave e o nome fornecidos.
'   Function gfRegistry_Write   Grava um valor na chave/nome fornecido,
'                       remove o nome se o valor passado for igual null [""]
'   Function gfRegistry_Create  Cria uma chave, Abre uma chave existente, podendo manter aberta ou criar e em segM46V999.gUida fecha-la(parametro de entrada).
'                       Esta função é chamada pela função gfRegistry_Write Se a chave não existir.
'                       Esta função só deve ser chamada diretamente, em casos muito específicos.
'   Function gfRegistry_DeleteKey Deleta a chave especificada e todas suas subchaves existentes
'                        A função mfRegistry_DeleteTree é sub-rotina exclusiva desta função
'                           (para excluir as subpastas)
'   Procedure gpRegistry_ListKey Lista todas as subchaves abaixo da chave especificada no parametro de entrada.
'   Procedure gpRegistry_AllowPermission Delega privilégios ilimitados para todos os usuários
'                                   na chave especificada no parametro de entrada.
'                                   Obs.: O usuário que executar esta função precisa ter direitos de gravação na chave especificada.
'                                   Nota: Esta função é chamada pela função gfRegistry_Create quando nesta a chave for recém criada
'                                   (É necessário pois ao criar uma chave com esta função, os direitos de gravação ficaram limitados oa usuario atual)
'                             As demais funções/Procedures, são sub chamadas exclusivas desta procedure
'                                   (mfConstructAndAddAce; mfIsEqual_SID
'                                    mfRegistry_AddSecurityDescriptor;
'                                    mfRegistry_UpdatePermissions)
'
' Obs:!!! Apenas dados dos tipos [REG_SZ] e [REG_EXPAND_SZ]
' estão homologados neste módulo (porém os demais podem funcionar)
'
' O módulos faz referencia a procedure gM46V999.gpGraLog (não declarada aqui!!!)
'
'-------------------------------------------------------------------------------------------

DefLng A-Z

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004

'Direitos de acessos genéricos
Public Const GENERIC_ALL = &H10000000
Public Const GENERIC_READ = &H80000000
Public Const GENERIC_EXECUTE = &H20000000
Public Const GENERIC_WRITE = &H40000000

'Dirteitos de acesso padrão
Public Const WRITE_DAC = &H40000

' Tipo de "ACE" sendo adicionado.
Public Const ACCESS_ALLOWED_ACE_TYPE = 0
Public Const ACCESS_DENIED_ACE_TYPE = 1

' Constantes do WINNT.H para os diversos conhecidos "SIDs", "users" e "groups"
Public Const SECURITY_WORLD_SID_AUTHORITY = &H1
Public Const SECURITY_WORLD_RID = &H0

' Flags de herança preenchidos no "AceFlags"
' Campo de um "Ace header".
'    Public Const OBJECT_INHERIT_ACE = &H1
Public Const CONTAINER_INHERIT_ACE = &H2
'    Public Const NO_PROPAGATE_INHERIT_ACE = &H4
Public Const INHERIT_ONLY_ACE = &H8
Public Const INHERITED_ACE = &H10
'    Public Const VALID_INHERIT_FLAGS = &H1F


' Canstantes utilizadas em call de API .
'Consulte MSDN para mais informações
'sobre o que?/como? estas constantes podem ser usadas.
Private Const DACL_SECURITY_INFORMATION = &H4
Private Const SECURITY_DESCRIPTOR_MIN_LENGTH = 20
Private Const SECURITY_DESCRIPTOR_REVISION = 1
Private Const ACL_REVISION = 2
Private Const MAXDWORD = &HFFFFFFFF

' Flags para "security descriptor".
Private Const SE_DACL_AUTO_INHERIT_REQ = &H100
'Private Const SE_SACL_AUTO_INHERIT_REQ = &H200
Private Const SE_DACL_AUTO_INHERITED = &H400
'Private Const SE_SACL_AUTO_INHERITED = &H800
Private Const SE_DACL_PROTECTED = &H1000
'Private Const SE_SACL_PROTECTED = &H2000

'Constantes de memória utilizadas por diversas "memory API calls".
Private Const LMEM_FIXED = &H0
Private Const LMEM_ZEROINIT = &H40
Private Const LPTR = (LMEM_FIXED + LMEM_ZEROINIT)

Const ERROR_FILE_NOT_FOUND& = 2
Const ERROR_BADKEY& = 1010
Const ERROR_MORE_DATA = 234    '  dderror
Const ERROR_NO_MORE_ITEMS = 259&
Public Const ERROR_SUCCESS& = 0
Public Const NO_ERROR& = 0
Const ACCESS_SYSTEM_SECURITY = &H1000000

Public Const KEY_ALL_ACCESS = &H3F
Const KEY_QUERY_VALUE = &H1&
Const KEY_SET_VALUE = &H2&
Const KEY_CREATE_SUB_KEY = &H4&
Const KEY_ENUMERATE_SUB_KEYS = &H8&
Const KEY_NOTIFY = &H10&
Const READ_CONTROL = &H20000
Const SYNCHRONIZE = &H100000
Const STANDARD_RIGHTS_READ = READ_CONTROL
Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY
Const REG_OPTION_NON_VOLATILE = 0&
Public Const REG_CREATED_NEW_KEY& = 1
Public Const REG_OPENED_EXISTING_KEY& = 2

' Constante de informação de versão
Private Const VER_PLATFORM_WIN32_NT = &H2

'Tipo ref. a informação de versão
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

'Tipos de valores no registro
Public Enum e_KeyType
    REG_NONE = 0                      'No value type
    REG_SZ = 1                        'Unicode nul terminated string
    REG_EXPAND_SZ = 2                 'Unicode nul terminated string
    REG_BINARY = 3                    'Free form binary
    REG_DWORD = 4                     '32-bit number
    REG_DWORD_LITTLE_ENDIAN = 4       '32-bit number (igual a  REG_DWORD)
    REG_DWORD_BIG_ENDIAN = 5          '32-bit number
    REG_LINK = 6                      'Symbolic Link (unicode)
    REG_MULTI_SZ = 7                  'Multiple Unicode strings
    REG_RESOURCE_LIST = 8             'Resource list in the resource map
    REG_FULL_RESOURCE_DESCRIPTOR = 9  'Resource list in the hardware description
End Enum

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

'Atributos de segurança
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescription As Long   'SECURITY_DESCRIPTOR
    bInheritHandle As Boolean
End Type

' Permissões de conta
Public Type AccountPerm
    AccountName As String
    AccessMask As Long
    AceFlags As Byte
    AceType As Byte
    pSid As Long
    SidPassedByCaller As Boolean
End Type

Private Type SDMemInfo
    pSD As Long
    pAcl As Long
End Type

Private Type SID_IDENTIFIER_AUTHORITY
    Value(6) As Byte
End Type

' Tipos para manipulação de "ACL".
' Consulte MSDN para maiores informações.
Private Type ACL
    AclRevision As Byte
    Sbz1 As Byte
    AclSize As Integer
    AceCount As Integer
    Sbz2 As Integer
End Type

Private Type ACL_SIZE_INFORMATION
    AceCount As Long
    AclBytesInUse As Long
    AclBytesFree As Long
End Type

Private Type ACE_HEADER
    AceType As Byte
    AceFlags As Byte
    AceSize As Integer
End Type

Private Type ACE
    Header As ACE_HEADER
    Mask As Long
    SidStart As Long
End Type

Private Declare Function GetVersionExA Lib "kernel32.dll" _
                                       (lpVersionInformation As OSVERSIONINFO) As Integer

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Declare Function RelCreateKeyEx Lib "advapi32.dll" Alias "RelCreateKeyExA" _
                                        (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
                                         ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, _
                                         lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, _
                                         lpdwDisposition As Long) As Long

Private Declare Function RelDeleteKey Lib "advapi32.dll" Alias "RelDeleteKeyA" _
                                      (ByVal hKey As Long, ByVal lpSubKey As String) As Long

Private Declare Function RelDeleteValue Lib "advapi32.dll" Alias "RelDeleteValueA" _
                                        (ByVal hKey As Long, ByVal lpValueName As String) As Long

Private Declare Function RegOpenKey Lib "advapi32" Alias "RegOpenKeyA" _
                                    (ByVal hKey As Long, ByVal lpValueName As String, phkResult As Long) As Long

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" _
                                      (ByVal hKey As Long, ByVal lpValueName As String, ByVal ulOptions As Long, _
                                       ByVal samDesired As Long, phkResult As Long) As Long

Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" _
                                         (ByVal hKey As Long, ByVal lpValueName As String, lpReserved As Long, _
                                          lpType As Long, ByVal lpData As String, lpcbData As Long) As Long

Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" _
                                       (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
                                        ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" _
                                    (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, _
                                     ByVal cbName As Long) As Long

Private Declare Function AllocateAndInitializeSid Lib "advapi32.dll" _
                                                  (pIdentifierAuthority As SID_IDENTIFIER_AUTHORITY, _
                                                   ByVal nSubAuthorityCount As Byte, ByVal nSubAuthority0 As Long, _
                                                   ByVal nSubAuthority1 As Long, ByVal nSubAuthority2 As Long, _
                                                   ByVal nSubAuthority3 As Long, ByVal nSubAuthority4 As Long, _
                                                   ByVal nSubAuthority5 As Long, ByVal nSubAuthority6 As Long, _
                                                   ByVal nSubAuthority7 As Long, lpPSid As Long) As Long

Private Declare Sub FreeSid Lib "advapi32.dll" (ByVal pSid As Long)

Private Declare Function RegGetKeySecurity Lib "advapi32.dll" _
                                           (ByVal hKey As Long, _
                                            ByVal SecurityInformation As Long, _
                                            ByVal pSecurityDescriptor As Long, _
                                            lpcbSecurityDescriptor As Long) As Long

Private Declare Function RegSetKeySecurity Lib "advapi32.dll" _
                                           (ByVal hKey As Long, _
                                            ByVal SecurityInformation As Long, _
                                            ByVal pSecurityDescriptor As Long) As Long

Private Declare Function LocalAlloc Lib "kernel32.dll" _
                                    (ByVal wFlags As Long, ByVal wBytes As Long) As Long

Private Declare Function LocalFree Lib "kernel32.dll" _
                                   (ByVal hMem As Long) As Long

Private Declare Function InitializeSecurityDescriptor Lib "advapi32.dll" _
                                                      (ByVal pSecurityDescriptor As Long, _
                                                       ByVal dwRevision As Long) As Long

Private Declare Function LookupAccountName Lib "advapi32.dll" Alias _
                                           "LookupAccountNameA" (ByVal lpSystemName As Long, _
                                                                 ByVal lpAccountName As String, _
                                                                 ByVal Sid As Long, _
                                                                 cbSid As Long, _
                                                                 ByVal ReferencedDomainName As String, _
                                                                 cbReferencedDomainName As Long, _
                                                                 peUse As Long) As Long

Private Declare Function GetSecurityDescriptorDacl Lib "advapi32.dll" _
                                                   (ByVal pSecurityDescriptor As Long, lpbDaclPresent As Long, _
                                                    pDacl As Long, lpbDaclDefaulted As Long) As Long

Private Declare Function GetAclInformation Lib "advapi32.dll" _
                                           (ByVal pAcl As Long, pAclInformation As Any, _
                                            ByVal nAclInformationLength As Long, _
                                            ByVal dwAclInformationClass As Long) As Long

Private Declare Function GetLengthSid Lib "advapi32.dll" _
                                      (ByVal pSid As Long) As Long

Private Declare Function InitializeAcl Lib "advapi32.dll" _
                                       (ByVal pAcl As Long, ByVal nAclLength As Long, _
                                        ByVal dwAclRevision As Long) As Long

Private Declare Function GetAce Lib "advapi32.dll" _
                                (ByVal pAcl As Long, ByVal dwAceIndex As Long, pACE As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" _
                               (hpvDest As Any, ByVal hpvSource As Long, _
                                ByVal cbCopy As Long)

Private Declare Function AddAce Lib "advapi32.dll" (ByVal pAcl As Long, _
                                                    ByVal dwAceRevision As Long, ByVal dwStartingAceIndex As Long, _
                                                    ByVal pAceList As Long, ByVal nAceListLength As Long) As Long

Private Declare Function GetSecurityDescriptorControl Lib "advapi32.dll" _
                                                      (ByVal pSecurityDescriptor As Long, _
                                                       pControl As Long, lpdwRevision As Long) As Long

Private Declare Function SetSecurityDescriptorControl Lib "advapi32.dll" _
                                                      (ByVal pSecurityDescriptor As Long, _
                                                       ByVal controlBitsOfInterest As Long, _
                                                       ByVal controlBitsToSet As Long) As Long

Private Declare Function SetSecurityDescriptorDacl Lib "advapi32.dll" _
                                                   (ByVal pSecurityDescriptor As Long, ByVal bDaclPresent As Long, _
                                                    ByVal pDacl As Long, ByVal bDaclDefaulted As Long) As Long

Private Declare Function EqualSid Lib "advapi32.dll" _
                                  (ByVal pSid1 As Long, ByVal pSid2 As Long) As Long

Public Enum E_AplMicroSoft
    E_AplMS_Word = 1
    E_AplMS_Excel = 2
    E_AplMS_PowerPoint = 3
End Enum


Public Function gfRegistry_Create(sKey As String, _
                                  ByVal pKeepOpen As Boolean, _
                                  ByRef pHkey As Long, _
                                  ByRef pErrMsg As String, _
                                  Optional vntOptionalHKey As Variant) As Long
'Criar a chave especificada
' Saída:
'    False se falhou na criação da chave,
'   lDisposition:
'    REG_CREATED_NEW_KEY& = 1& or   'Nova Chave criada
'    REG_OPENED_EXISTING_KEY& = 2&  'Abriu uma chave existente
'
    Dim lRtn As Integer
    Dim lHKey As Long                 'handle da chave aberta
    Dim lDisposition As Long    'disposition
    Dim lpSecurityAttributes As SECURITY_ATTRIBUTES
    Dim lOptionalHKey As Long

    pErrMsg = ""
    If IsMissing(vntOptionalHKey) Then
        lOptionalHKey = HKEY_LOCAL_MACHINE   'Utiliza local HKLM
    Else
        lOptionalHKey = vntOptionalHKey     'Utiliza a chave informada
    End If

Executa:
    lRtn = RelCreateKeyEx(lOptionalHKey, sKey, 0&, "", _
                          REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, _
                          lpSecurityAttributes, _
                          lHKey, lDisposition)

    If lRtn = ERROR_SUCCESS Then
        gfRegistry_Create = lDisposition    'Informa se a chave foi criada ou apenas aberta (caso já existente)
        If pKeepOpen = False Then
            lRtn = RegCloseKey(lHKey)    'close the Registry
            'Altera as propriedades para permitir que outros
            'usuarios tenham permissão sobre o registro criado
            Call gpRegistry_AllowPermission(sKey, lOptionalHKey)
        Else
            If lDisposition = REG_CREATED_NEW_KEY Then
                lRtn = RegCloseKey(lHKey)
                Call gpRegistry_AllowPermission(sKey, lOptionalHKey)
                GoTo Executa
            End If
            gfRegistry_Create = lRtn    'Retorno da criação da chave
            pHkey = lHKey
        End If
    Else
        If pKeepOpen = True Then
            gfRegistry_Create = lRtn    'Retorno da criação da chave
            pHkey = lHKey
        End If
        GoTo Func_Erro
    End If

Func_Saida:
    Exit Function

Func_Erro:
    Select Case lRtn
    Case 0      'ERROR_NONE = 0
        pErrMsg = "Sem erro"
    Case 1      'ERROR_BADDB = 1
        pErrMsg = "ERROR_BADDB"
    Case 2      'ERROR_BADKEY = 2
        pErrMsg = "ERROR_BADKEY"
    Case 3      'ERROR_CANTOPEN = 3
        pErrMsg = "Erro ao abrir registro. (Acesso negado)"
    Case 4      'ERROR_CANTREAD = 4
        pErrMsg = "Erro ao ler registro. (Acesso negado)"
    Case 5      'ERROR_CANTWRITE = 5
        pErrMsg = "Erro ao gravar registro. (Acesso negado)"
    Case 6      'ERROR_OUTOFMEMORY = 6
        pErrMsg = "ERROR_OUTOFMEMORY"
    Case 7      'ERROR_INVALID_PARAMETER = 7
        pErrMsg = "ERROR_INVALID_PARAMETER"
    Case 8      'ERROR_ACCESS_DENIED = 8
        pErrMsg = "ERROR_ACCESS_DENIED"
    Case 87     'ERROR_INVALID_PARAMETERS = 87
        pErrMsg = "ERROR_INVALID_PARAMETERS"
    Case 259    'ERROR_NO_MORE_ITEMS = 259
        pErrMsg = "ERROR_NO_MORE_ITEMS"
    Case Else
        pErrMsg = "Erro desconhecido.(" & lRtn & ")"
    End Select
    pErrMsg = "gfRegistry_Create - " & pErrMsg & Chr(13) & sKey

    GoTo Func_Saida
End Function

Public Function gfRegistry_Read(sKey As String, sValueName As String, pKeyType As e_KeyType, _
                                Optional vntOptionalHKey As Variant) As String
' Retorna o valor econtrada para sKey and sValueName
' Entrada:        Exemplo:
'   sKey        "Software\Microsoft\File Manager\Settings"
'   sValueName  "Face"
' Saida:
'               "FixedSys" or
'               "" [null] if not found
'-----------------------------------------------------------------------------------

    Dim lOptionalHKey As Long
    Dim lHKey As Long
    Dim lpcbData As Long    'tamanho da string retornada
    Dim sReturnedString As String
    Dim sTemp As String
    Dim lRtn As Long

    If IsMissing(vntOptionalHKey) Then
        lOptionalHKey = HKEY_LOCAL_MACHINE
    Else
        lOptionalHKey = vntOptionalHKey
    End If

    lRtn = RegOpenKeyEx(lOptionalHKey, sKey, 0&, KEY_READ, lHKey)
    If lRtn = ERROR_SUCCESS Then
        lpcbData = 1024                     'Prevê muitos caracteres
        sReturnedString = Space$(lpcbData)  'Configura o Buffer
        lRtn = RegQueryValueEx(lHKey, sValueName, ByVal 0&, pKeyType, sReturnedString, lpcbData)
        If lRtn = ERROR_SUCCESS Then
            sTemp = Left$(sReturnedString, lpcbData - 1)
        End If
        RegCloseKey lHKey
    End If
    gfRegistry_Read = sTemp
End Function

Public Function gfRegistry_Write(sKey As String, sValueName As String, sValue As String, _
                                 pKeyType As e_KeyType, _
                                 ByRef pErrMsg As String, _
                                 Optional vntOptionalHKey As Variant) As Integer
' Entrada:      Exemplo:
'   sKey        "Software\Microsoft\File Manager\Settings"
'   sValueName  "Face"
'   sValue      "FixedSys"
' Saida:
'   True caso OK!
'
' Note: Se sValue = "" então sValueName é removido [deletado].
'-----------------------------------------------------------------------------------
    Dim lOptionalHKey As Long
    Dim lRtn As Long
    Dim lHKey As Long
    Dim iSuccessCount As Integer
    Dim lChaveAtual As String

    pErrMsg = ""

    If IsMissing(vntOptionalHKey) Then
        lOptionalHKey = HKEY_LOCAL_MACHINE
    Else
        lOptionalHKey = vntOptionalHKey
    End If

    If Trim$(sValue) <> "" Then
RegWriteTryAgain:
        'Abre/Cria a chave para atualização
        lRtn = gfRegistry_Create(sKey, True, lHKey, pErrMsg, lOptionalHKey)
        lChaveAtual = sKey
        If lRtn = ERROR_SUCCESS Then
            lRtn = RegSetValueEx(lHKey, sValueName, 0&, pKeyType, ByVal sValue, CLng(Len(sValue) + 1))    'update the value
            lChaveAtual = sValueName
            If lRtn = ERROR_SUCCESS Then
                iSuccessCount = iSuccessCount + 1
            End If
            lRtn = RegCloseKey(lHKey)
        ElseIf lRtn = ERROR_FILE_NOT_FOUND Or lRtn = ERROR_BADKEY Then    'Cria uma
            If gfRegistry_Create(sKey, False, lHKey, pErrMsg, lOptionalHKey) Then
                'Criado com sucesso
                'Tenta gravar o registro novamente
                GoTo RegWriteTryAgain
            End If
        End If
    Else
        'Valor nulo, exclui a chave
        lRtn = RegOpenKeyEx(lOptionalHKey, sKey, 0&, KEY_ALL_ACCESS, lHKey)  'open the Registry for update
        lChaveAtual = sKey
        If lRtn = ERROR_SUCCESS Then
            lRtn = RelDeleteValue(lHKey, sValueName)
            lChaveAtual = sValueName
            If lRtn = ERROR_SUCCESS Then
                iSuccessCount = iSuccessCount + 1
            End If
            lRtn = RegCloseKey(lHKey)
        End If
    End If
    If iSuccessCount > 0 Then
        'OK!!!
        gfRegistry_Write = True
    Else
        GoTo Func_Erro
    End If

Func_Saida:
    Exit Function

Func_Erro:
    Select Case lRtn
    Case 0      'ERROR_NONE = 0
        pErrMsg = "Sem erro"
    Case 1      'ERROR_BADDB = 1
        pErrMsg = "ERROR_BADDB"
    Case 2      'ERROR_BADKEY = 2
        pErrMsg = "ERROR_BADKEY"
    Case 3      'ERROR_CANTOPEN = 3
        pErrMsg = "Erro ao abrir registro. (Acesso negado)"
    Case 4      'ERROR_CANTREAD = 4
        pErrMsg = "Erro ao ler registro. (Acesso negado)"
    Case 5      'ERROR_CANTWRITE = 5
        pErrMsg = "Erro ao gravar registro. (Acesso negado)"
    Case 6      'ERROR_OUTOFMEMORY = 6
        pErrMsg = "ERROR_OUTOFMEMORY"
    Case 7      'ERROR_INVALID_PARAMETER = 7
        pErrMsg = "ERROR_INVALID_PARAMETER"
    Case 8      'ERROR_ACCESS_DENIED = 8
        pErrMsg = "ERROR_ACCESS_DENIED"
    Case 87     'ERROR_INVALID_PARAMETERS = 87
        pErrMsg = "ERROR_INVALID_PARAMETERS"
    Case 259    'ERROR_NO_MORE_ITEMS = 259
        pErrMsg = "ERROR_NO_MORE_ITEMS"
    Case Else
        pErrMsg = "Erro desconhecido.(" & lRtn & ")"
    End Select
    pErrMsg = "gfRegistry_Write - " & pErrMsg & Chr(13) & lChaveAtual

    GoTo Func_Saida

End Function

Public Function gfRegistry_DeleteKey(pStartKey As Variant, _
                                     pKeyName As String, _
                                     ByRef pErrMsg As String, _
                                     Optional vntOptionalHKey As Variant) As Integer

'*********!!!!!%%%%%%%%%%%%%%%%%
'       EXTREMO CUIDADO COM ESTA FUNÇÃO,
'       PAMARETROS DE ENTRADA ENVIADOS, NÃO INTENCIONALMENTE
'       PODE AFETAR O FUNCIONAMENTO DE OUTROS PROGRAMAS
'       OU ATÉ DANIFICAR O SISTEMA OPERACIONAL
'
'***********%%%%%%%%%%%%%%%%%%

    Dim lOptionalHKey As Long
    Dim lRtn As Long
    Dim lHKey As Long
    Dim iSuccessCount As Integer
    Dim lretorno As String
    Dim lChaveAtual As String

    pErrMsg = ""
    If IsMissing(vntOptionalHKey) Then
        lOptionalHKey = HKEY_LOCAL_MACHINE
    Else
        lOptionalHKey = vntOptionalHKey
    End If

    If pStartKey = "" Or _
       pKeyName = "" Then
        Exit Function
    End If

    'Deleta as subchaves existentes
    lretorno = mfRegistry_DeleteTree(vntOptionalHKey, pStartKey & "\" & pKeyName, pErrMsg)
    If Val(lretorno) <> ERROR_SUCCESS And _
       Val(lretorno) <> 2 Then    'Excluido com sucesso / Diretório não existe
        'Não excluiu SubChaves
        Exit Function
    End If

    'Deleta a Chave atual
    lRtn = RegOpenKeyEx(lOptionalHKey, pStartKey, 0&, KEY_SET_VALUE, lHKey)  'open the Registry for update
    lChaveAtual = pStartKey
    If lRtn = ERROR_SUCCESS Then
        lRtn = RelDeleteKey(lHKey, pKeyName)
        lChaveAtual = pKeyName
        If lRtn = ERROR_SUCCESS Or _
           lRtn = 2 Then         'Excluido com sucesso / Diretório não existe
            iSuccessCount = iSuccessCount + 1
        End If
        lRtn = RegCloseKey(lHKey)
    End If

    If iSuccessCount > 0 Then
        'OK!!
        gfRegistry_DeleteKey = True
    Else
        GoTo Func_Erro
    End If

Func_Saida:
    Exit Function

Func_Erro:
    Select Case lRtn
    Case 0      'ERROR_NONE = 0
        pErrMsg = "Sem erro"
    Case 1      'ERROR_BADDB = 1
        pErrMsg = "ERROR_BADDB"
    Case 2      'ERROR_BADKEY = 2
        pErrMsg = "ERROR_BADKEY"
    Case 3      'ERROR_CANTOPEN = 3
        pErrMsg = "Erro ao abrir registro. (Acesso negado)"
    Case 4      'ERROR_CANTREAD = 4
        pErrMsg = "Erro ao ler registro. (Acesso negado)"
    Case 5      'ERROR_CANTWRITE = 5
        pErrMsg = "Erro ao gravar registro. (Acesso negado)"
    Case 6      'ERROR_OUTOFMEMORY = 6
        pErrMsg = "ERROR_OUTOFMEMORY"
    Case 7      'ERROR_INVALID_PARAMETER = 7
        pErrMsg = "ERROR_INVALID_PARAMETER"
    Case 8      'ERROR_ACCESS_DENIED = 8
        pErrMsg = "ERROR_ACCESS_DENIED"
    Case 87     'ERROR_INVALID_PARAMETERS = 87
        pErrMsg = "ERROR_INVALID_PARAMETERS"
    Case 259    'ERROR_NO_MORE_ITEMS = 259
        pErrMsg = "ERROR_NO_MORE_ITEMS"
    Case Else
        pErrMsg = "Erro desconhecido.(" & lRtn & ")"
    End Select
    pErrMsg = "gfRegistry_DeleteKey - " & pErrMsg & Chr(13) & lChaveAtual

    GoTo Func_Saida

End Function

Private Function mfRegistry_DeleteTree(pStartKey As Variant, _
                                       pKeyName As String, _
                                       ByRef pErrMsg As String) As String

    Dim dwRtn As Long
    Dim dwSubKeyLength As Long
    Dim szSubKey As String
    Dim hKey As Long
    Dim lCount As Long
    Dim lColSubkeys As Collection
    Dim lChaveAtual As String

    pErrMsg = ""
    On Error GoTo Func_Saida
    If pKeyName <> "" Then
        dwRtn = RegOpenKeyEx(pStartKey, pKeyName, 0&, _
                             KEY_ENUMERATE_SUB_KEYS, hKey)
        lChaveAtual = pKeyName
        If dwRtn = ERROR_SUCCESS Then
            Set lColSubkeys = New Collection
            lCount = 0
            Do
                ' Percorre as subkeys até ocorrer um erro
                dwSubKeyLength = 256
                szSubKey = Space$(dwSubKeyLength)
                If RegEnumKey(hKey, lCount, szSubKey, dwSubKeyLength) _
                   <> ERROR_SUCCESS Then Exit Do
                lCount = lCount + 1

                szSubKey = Left$(szSubKey, InStr(szSubKey, Chr$(0)) - 1)
                lColSubkeys.Add szSubKey
            Loop

            'Recursivamente deleta as subkeys e suas subkeys...
            For lCount = 1 To lColSubkeys.Count
                'Exclui as subkeys da subkeys atual.
                mfRegistry_DeleteTree pStartKey, pKeyName & "\" & lColSubkeys(lCount), pErrMsg

                'Deleta a subkeys atual
                RelDeleteKey hKey, lColSubkeys(lCount)
            Next lCount

            RegCloseKey (hKey)
        Else
            GoTo Func_Erro
        End If
    Else
        dwRtn = ERROR_BADKEY
    End If
Func_Saida:
    mfRegistry_DeleteTree = dwRtn
    Err = 0
    Exit Function

Func_Erro:
    Select Case dwRtn
    Case 0      'ERROR_NONE = 0
        pErrMsg = "Sem erro"
    Case 1      'ERROR_BADDB = 1
        pErrMsg = "ERROR_BADDB"
    Case 2      'ERROR_BADKEY = 2
        pErrMsg = "ERROR_BADKEY"
    Case 3      'ERROR_CANTOPEN = 3
        pErrMsg = "Erro ao abrir registro. (Acesso negado)"
    Case 4      'ERROR_CANTREAD = 4
        pErrMsg = "Erro ao ler registro. (Acesso negado)"
    Case 5      'ERROR_CANTWRITE = 5
        pErrMsg = "Erro ao gravar registro. (Acesso negado)"
    Case 6      'ERROR_OUTOFMEMORY = 6
        pErrMsg = "ERROR_OUTOFMEMORY"
    Case 7      'ERROR_INVALID_PARAMETER = 7
        pErrMsg = "ERROR_INVALID_PARAMETER"
    Case 8      'ERROR_ACCESS_DENIED = 8
        pErrMsg = "ERROR_ACCESS_DENIED"
    Case 87     'ERROR_INVALID_PARAMETERS = 87
        pErrMsg = "ERROR_INVALID_PARAMETERS"
    Case 259    'ERROR_NO_MORE_ITEMS = 259
        pErrMsg = "ERROR_NO_MORE_ITEMS"
    Case Else
        pErrMsg = "Erro desconhecido.(" & dwRtn & ")"
    End Select
    pErrMsg = "mfRegistry_DeleteTree - " & pErrMsg & Chr(13) & lChaveAtual

    GoTo Func_Saida

End Function

Public Sub gpRegistry_ListKey(pStartKey As Variant, _
                              pKeyName As String, _
                              ByVal pPrefix As String, _
                              ByRef pRetList As String)
'Retorna lista de Chaves com o prefixo informado,
'que esteja imetdiatamente abaixo da chave de parametro

    Dim dwRtn As Long
    Dim dwSubKeyLength As Long
    Dim szSubKey As String
    Dim hKey As Long
    Dim lCount As Long

    On Error GoTo Proc_SAIDA
    pRetList = ""
    If pKeyName <> "" Then
        dwRtn = RegOpenKeyEx(pStartKey, pKeyName, 0&, _
                             KEY_ENUMERATE_SUB_KEYS, hKey)
        If dwRtn = ERROR_SUCCESS Then
            Do
                ' Percorre as subkeys até ocorrer um erro
                dwSubKeyLength = 256
                szSubKey = Space$(dwSubKeyLength)
                If RegEnumKey(hKey, lCount, szSubKey, dwSubKeyLength) _
                   <> ERROR_SUCCESS Then Exit Do

                lCount = lCount + 1

                szSubKey = Left$(szSubKey, InStr(szSubKey, Chr$(0)) - 1)
                If UCase(Left(szSubKey, Len(pPrefix))) = UCase(pPrefix) Then
                    pRetList = pRetList & IIf(pRetList = "", "", ";")
                    pRetList = pRetList & szSubKey
                End If
            Loop

            RegCloseKey (hKey)

        End If
    Else
        dwRtn = ERROR_BADKEY
    End If


Proc_SAIDA:
    Err = 0
    Exit Sub

End Sub

Public Sub gpRegistry_AllowPermission(ByVal pKeyName As String, _
                                      Optional pOptionalHKey As Variant)

    Dim lOptionalHKey As Long
    Dim hKey As Long
    Dim Accounts(0 To 1) As AccountPerm
    Dim fResult As Long, n As Long
    Dim dwNumOfAccounts As Long
    Dim siaNtAuthority As SID_IDENTIFIER_AUTHORITY

    If IsMissing(pOptionalHKey) Then
        lOptionalHKey = HKEY_LOCAL_MACHINE
    Else
        lOptionalHKey = pOptionalHKey
    End If

    dwNumOfAccounts = UBound(Accounts)

    'Configura, em um array, as permissões de contas para todos os usuarios

    '-- Permissão de leitura
    Accounts(0).AccountName = ""          'Todos os usuários
    Accounts(0).AccessMask = GENERIC_READ
    Accounts(0).AceFlags = CONTAINER_INHERIT_ACE
    Accounts(0).AceType = ACCESS_ALLOWED_ACE_TYPE
    'Constrói um SID para todos os usuários "Universal well-known SID"
    siaNtAuthority.Value(5) = SECURITY_WORLD_SID_AUTHORITY
    If AllocateAndInitializeSid(siaNtAuthority, 1, _
                                SECURITY_WORLD_RID, 0, 0, 0, 0, 0, 0, 0, _
                                Accounts(0).pSid) = 0 Then
        'Call gM46V999.gpGraLog(0, "gpRegistry_AllowPermission Erro - AllocateAndInitializeSid failed with error code : " _
         '              & Err.LastDllError)
        Exit Sub
    End If
    'If the caller initializes SID, set SidPassedByCaller member to True
    Accounts(0).SidPassedByCaller = True

    '-- Permissão Total na chave
    Accounts(1).AccountName = ""        'Todos os usuários
    Accounts(1).AccessMask = GENERIC_ALL
    Accounts(1).AceFlags = 0
    Accounts(1).AceType = ACCESS_ALLOWED_ACE_TYPE
    'Constrói um SID para todos os usuários "Universal well-known SID"
    siaNtAuthority.Value(5) = SECURITY_WORLD_SID_AUTHORITY
    If AllocateAndInitializeSid(siaNtAuthority, 1, _
                                SECURITY_WORLD_RID, 0, 0, 0, 0, 0, 0, 0, _
                                Accounts(1).pSid) = 0 Then
        'Call gM46V999.gpGraLog(0, "gpRegistry_AllowPermission Erro - AllocateAndInitializeSid failed with error code : " _
         '              & Err.LastDllError)
        Exit Sub
    End If
    Accounts(1).SidPassedByCaller = False


    fResult = RegOpenKeyEx(lOptionalHKey, _
                           pKeyName, 0, _
                           READ_CONTROL Or WRITE_DAC, hKey)
    If fResult <> ERROR_SUCCESS Then
        'Call gM46V999.gpGraLog(0, "gpRegistry_AllowPermission Erro - RegOpenKeyEx failed with error code : " & fResult)
    Else
        mfRegistry_UpdatePermissions hKey, Accounts
        RegCloseKey hKey
    End If

    'Limpa da memória
    For n = 0 To dwNumOfAccounts
        If Accounts(n).pSid <> 0 And Accounts(n).SidPassedByCaller Then
            FreeSid (Accounts(n).pSid)
            Accounts(n).pSid = 0
        End If
    Next
End Sub

Private Function mfRegistry_UpdatePermissions(ByVal hKey As Long, _
                                              Accounts() As AccountPerm) As Boolean

    Dim fResult As Long
    Dim sdInfo As SDMemInfo
    Dim oldSD As Long
    Dim nLengthNeeded As Long
    Dim bStatus As Boolean

    bStatus = False
    On Error GoTo Cleanup

    sdInfo.pAcl = 0
    sdInfo.pSD = 0

    nLengthNeeded = 0
    fResult = RegGetKeySecurity(hKey, _
                                DACL_SECURITY_INFORMATION, 0, _
                                nLengthNeeded)
    ' This call will fail. On Return nLengthNeeded will be updated.
    ' Check for that below
    If nLengthNeeded = 0 Then
        'Call gM46V999.gpGraLog(0, "mfRegistry_UpdatePermissions Erro - RegGetKeySecurity failed with error code : " _
         '             & fResult)
        Err.Raise 0
    End If

    oldSD = LocalAlloc(LPTR, nLengthNeeded)
    If oldSD = 0 Then
        'Call gM46V999.gpGraLog(0, "mfRegistry_UpdatePermissions Erro - LocalAlloc failed with error code : " _
         '             & Err.LastDllError)
        Err.Raise 0
    End If

    fResult = RegGetKeySecurity(hKey, _
                                DACL_SECURITY_INFORMATION, oldSD, _
                                nLengthNeeded)
    If fResult <> ERROR_SUCCESS Then
        'Call gM46V999.gpGraLog(0, "mfRegistry_UpdatePermissions Erro - RegGetKeySecurity failed with error code : " _
         '              & fResult)
        Err.Raise 0
    End If

    fResult = mfRegistry_AddSecurityDescriptor(oldSD, Accounts(), sdInfo)
    If fResult = 0 Then
        'Call gM46V999.gpGraLog(0, "mfRegistry_UpdatePermissions Erro - Unable to create Security Descriptor")
        Err.Raise 0
    End If

    fResult = RegSetKeySecurity(hKey, _
                                DACL_SECURITY_INFORMATION, sdInfo.pSD)
    If fResult <> ERROR_SUCCESS Then
        'Call gM46V999.gpGraLog(0, "mfRegistry_UpdatePermissions Erro - RegSetKeySecurity failed with error code : " _
         '              & fResult)
        Err.Raise 0
    End If

    bStatus = True

Cleanup:
    'Free the memory allocated
    If (oldSD <> 0) Then LocalFree oldSD
    oldSD = 0
    If (sdInfo.pSD <> 0) Then LocalFree sdInfo.pSD
    sdInfo.pSD = 0
    If (sdInfo.pAcl <> 0) Then LocalFree sdInfo.pAcl
    sdInfo.pAcl = 0
    mfRegistry_UpdatePermissions = bStatus
End Function

Private Function mfRegistry_AddSecurityDescriptor(ByVal pOldSD As Long, _
                                                  Accounts() As AccountPerm, _
                                                  sdInfo As SDMemInfo) As Long

    Dim pNewACL As Long
    Dim dwNewACLSize As Long
    Dim dwTotalDACLSize As Long
    Dim szDomainName As String
    Dim cbDomainName As Long
    Dim nSidSize As Long
    Dim i As Long, n As Long
    Dim eUse As Long
    Dim fReturn As Long
    Dim fResult As Long
    Dim tempACL As ACL
    Dim tempAce As ACE
    Dim Ptr As Long
    Dim dwNumOfAccounts As Long
    Dim pSD As Long
    Dim AceIndex As Long
    Dim lDaclPresent As Long
    Dim lDaclDefaulted As Long
    Dim sACLInfo As ACL_SIZE_INFORMATION
    Dim pAcl As Long
    Dim osinfo As OSVERSIONINFO
    Dim w2kOrAbove As Boolean

    On Error GoTo ExitLabel

    ' Determine if system is Windows 2000 or above
    osinfo.dwOSVersionInfoSize = Len(osinfo)
    osinfo.szCSDVersion = Space$(128)
    GetVersionExA osinfo
    w2kOrAbove = _
    (osinfo.dwPlatformId = VER_PLATFORM_WIN32_NT And _
     osinfo.dwMajorVersion >= 5)

    ' Intialize some of the variables
    fReturn = 0
    sdInfo.pAcl = 0
    sdInfo.pSD = 0

    dwNumOfAccounts = UBound(Accounts)

    ' Allocate memory for a new Security Descriptor
    pSD = LocalAlloc(LPTR, SECURITY_DESCRIPTOR_MIN_LENGTH)
    If pSD = 0 Then Err.Raise 0
    sdInfo.pSD = pSD

    ' Initialize the new Security Descriptor
    fResult = InitializeSecurityDescriptor(pSD, SECURITY_DESCRIPTOR_REVISION)
    If fResult = 0 Then Err.Raise 0

    ' Get the existing ACL size
    lDaclPresent = 0
    pAcl = 0
    If (pOldSD) Then
        fResult = GetSecurityDescriptorDacl(pOldSD, lDaclPresent, _
                                            pAcl, lDaclDefaulted)
        If fResult = 0 Then Err.Raise 0

        If (lDaclPresent <> 0 And pAcl <> 0) Then
            fResult = GetAclInformation(pAcl, sACLInfo, Len(sACLInfo), 2&)
            If fResult = 0 Then Err.Raise 0
            dwTotalDACLSize = sACLInfo.AclBytesInUse
        Else
            dwTotalDACLSize = Len(tempACL)
        End If
    Else
        dwTotalDACLSize = Len(tempACL)
    End If

    ' Find the SIDs for each userName supplied in Accounts() array
    ' and compute the new ACL size needed.
    ' Call LookupAccountName only for the entries where the
    ' SID is not supplied by the caller.
    szDomainName = Space(256)
    For n = 0 To dwNumOfAccounts
        If (Accounts(n).pSid = 0) Then
            nSidSize = 0
            cbDomainName = 256

            ' Lookup the SID for this user
            ' First call is to find the buffer size required for SID
            fResult = LookupAccountName(0, Accounts(n).AccountName, 0, _
                                        nSidSize, szDomainName, _
                                        cbDomainName, eUse)
            Accounts(n).pSid = LocalAlloc(LPTR, nSidSize)
            If Accounts(n).pSid = 0 Then Err.Raise 0

            ' Get the Actual SID value in this second call
            fResult = LookupAccountName(0, Accounts(n).AccountName, _
                                        Accounts(n).pSid, _
                                        nSidSize, szDomainName, _
                                        cbDomainName, eUse)
            If fResult = 0 Then Err.Raise 0
        End If

        ' sizeof(DWORD) = 4
        dwNewACLSize = Len(tempAce) + GetLengthSid(Accounts(n).pSid) - 4
        dwTotalDACLSize = dwTotalDACLSize + dwNewACLSize
    Next

    ' Allocate memory for the new ACL
    pNewACL = LocalAlloc(LPTR, dwTotalDACLSize)
    If pNewACL = 0 Then Err.Raise 0

    sdInfo.pAcl = pNewACL

    ' Initialize the new ACL
    fResult = InitializeAcl(pNewACL, dwTotalDACLSize, ACL_REVISION)
    If fResult = 0 Then Err.Raise 0

    AceIndex = 0

    ' Add the new ACCESS DENIED ACEs first to the DACL
    For n = 0 To dwNumOfAccounts
        If (Accounts(n).AceType = ACCESS_DENIED_ACE_TYPE) Then
            fResult = mfConstructAndAddAce(pNewACL, _
                                           Accounts(n).AceType, _
                                           Accounts(n).AceFlags, _
                                           Accounts(n).AccessMask, _
                                           Accounts(n).pSid)
            If fResult = 0 Then Err.Raise 0
            AceIndex = AceIndex + 1
        End If
    Next

    ' Copy all non-inherited ACEs from the existing DACL
    If (lDaclPresent <> 0 And pAcl <> 0 And sACLInfo.AceCount > 0) Then
        ' Get each ACE from the old DACL and add them into the new DACL.
        For i = 0 To (sACLInfo.AceCount - 1)
            ' Attempt to get the next ACE.
            fResult = GetAce(pAcl, i, Ptr)
            If (fResult = 0) Then Err.Raise 0

            CopyMemory tempAce, Ptr, LenB(tempAce)
            ' Exit this for loop, once the first INHERITED_ACE is found
            If ((tempAce.Header.AceFlags And INHERITED_ACE) = INHERITED_ACE) Then
                Exit For
            End If

            'Add the ACE to the new DACL if the SID is not in Accounts()
            If Not (mfIsEqual_SID(Accounts(), Ptr + 8)) Then
                ' Now that you have the ACE, add it to the new ACL.
                fResult = AddAce(pNewACL, ACL_REVISION, _
                                 MAXDWORD, Ptr, _
                                 tempAce.Header.AceSize)
                If fResult = 0 Then Err.Raise 0
                AceIndex = AceIndex + 1
            End If
        Next i
    End If

    ' Add the new ACCESS ALLOWED ACEs next to the DACL
    For n = 0 To dwNumOfAccounts
        If (Accounts(n).AceType = ACCESS_ALLOWED_ACE_TYPE) Then
            fResult = mfConstructAndAddAce(pNewACL, _
                                           Accounts(n).AceType, _
                                           Accounts(n).AceFlags, _
                                           Accounts(n).AccessMask, _
                                           Accounts(n).pSid)
            If fResult = 0 Then Err.Raise 0
            AceIndex = AceIndex + 1
        End If
    Next

    ' Copy now all inherited ACEs from the existing DACL, so that the
    ' new DACL will be in the Windows 2000 preferred order
    If (lDaclPresent <> 0 And pAcl <> 0 And sACLInfo.AceCount > 0) Then
        ' Get each INHERITED_ACE from the old ACL and
        ' add them into the new ACL.
        For i = i To (sACLInfo.AceCount - 1)
            ' Attempt to get the next ACE.
            fResult = GetAce(pAcl, i, Ptr)
            If (fResult = 0) Then Err.Raise 0

            CopyMemory tempAce, Ptr, LenB(tempAce)
            ' Add it to the new ACL.
            fResult = AddAce(pNewACL, ACL_REVISION, _
                             MAXDWORD, Ptr, _
                             tempAce.Header.AceSize)
            If fResult = 0 Then Err.Raise 0
            AceIndex = AceIndex + 1
        Next i
    End If

    If w2kOrAbove And pOldSD <> 0 Then
        Dim controlFlag As Long
        Dim dwRevision As Long
        Dim controlBitsOfInterest As Long
        Dim controlBitsToSet As Long

        fResult = GetSecurityDescriptorControl(pOldSD, _
                                               controlFlag, dwRevision)
        If (fResult <> 0) Then
            controlBitsOfInterest = 0
            controlBitsToSet = 0
            If ((controlFlag And SE_DACL_AUTO_INHERITED) = _
                SE_DACL_AUTO_INHERITED) Then
                controlBitsOfInterest = _
                SE_DACL_AUTO_INHERIT_REQ Or _
                                        SE_DACL_AUTO_INHERITED
                controlBitsToSet = controlBitsOfInterest
            ElseIf ((controlFlag And SE_DACL_PROTECTED) = _
                    SE_DACL_PROTECTED) Then
                controlBitsOfInterest = _
                SE_DACL_PROTECTED
                controlBitsToSet = controlBitsOfInterest
            End If
            If controlBitsToSet <> 0 Then
                fResult = SetSecurityDescriptorControl(pSD, _
                                                       controlBitsOfInterest, _
                                                       controlBitsToSet)
                If fResult = 0 Then Err.Raise 0
            End If
        End If
    End If

    ' Add the new DACL to the new Security Descriptor
    fResult = SetSecurityDescriptorDacl(pSD, 1, pNewACL, 0)
    If fResult = 0 Then Err.Raise 0

    fReturn = 1

ExitLabel:
    ' Make sure we clean up
    For n = 0 To dwNumOfAccounts
        ' Free only the SIDs that has been allocated in this function
        If Accounts(n).pSid <> 0 And _
           Not (Accounts(n).SidPassedByCaller) Then
            LocalFree (Accounts(n).pSid)
            Accounts(n).pSid = 0
        End If
    Next
    ' If any of the functions failed, free new SD, and new ACL
    If fReturn = 0 Then
        If (sdInfo.pSD <> 0) Then LocalFree sdInfo.pSD
        sdInfo.pSD = 0
        If (sdInfo.pAcl <> 0) Then LocalFree sdInfo.pAcl
        sdInfo.pAcl = 0
    End If

    mfRegistry_AddSecurityDescriptor = fReturn

End Function

Private Function mfConstructAndAddAce(ByVal pNewACL As Long, _
                                      ByVal AceType As Byte, _
                                      ByVal AceFlags As Byte, _
                                      ByVal AccessMask As Long, _
                                      ByVal pSid As Long) As Long

    Dim fResult As Long
    Dim dwNewACESize As Long
    Dim dwSidLen As Long
    Dim tempAce As ACE
    Dim pACE As Long

    fResult = 0
    On Error GoTo Retorno

    ' Find the length of SID and size of new ACE to be added
    dwSidLen = GetLengthSid(pSid)
    dwNewACESize = Len(tempAce) + dwSidLen - 4

    ' Allocate memory for the new ACE
    pACE = LocalAlloc(LPTR, dwNewACESize)
    If pACE = 0 Then Err.Raise 0

    ' Set up the ACE structure in VB variable
    tempAce.Header.AceType = AceType
    tempAce.Header.AceFlags = AceFlags
    tempAce.Header.AceSize = dwNewACESize
    tempAce.Mask = AccessMask

    ' Copy the VB variable contents and the SID to the
    ' the ACE allocated
    CopyMemory ByVal pACE, VarPtr(tempAce), LenB(tempAce)
    CopyMemory ByVal pACE + 8, pSid, dwSidLen

    ' Add the new ACE to the new ACL
    fResult = AddAce(pNewACL, ACL_REVISION, _
                     MAXDWORD, _
                     pACE, _
                     dwNewACESize)
    LocalFree pACE
Retorno:
    mfConstructAndAddAce = fResult
End Function

Private Function mfIsEqual_SID(Accounts() As AccountPerm, pSid As Long) As Boolean
    Dim nEntries As Long
    Dim nIndex As Long

    ' Check if the supplied SID pSid matches with one of the
    ' new SIDs specified in Accounts()
    nEntries = UBound(Accounts)
    For nIndex = 0 To nEntries
        If (EqualSid(Accounts(nIndex).pSid, pSid)) Then
            mfIsEqual_SID = True
            Exit Function
        End If
    Next
    mfIsEqual_SID = False
End Function

Public Function gfPathAplMicrosoft(ByVal pAplMS As E_AplMicroSoft) As String
'Retorna o Path para executar os aplicativos Microsoft
'Independente da versao instalada...

    Dim lFullPath As String
    Dim lMsgErr As String
    Dim lChavesOffice As String
    Dim lVersao(5) As String
    Dim lAplicativo As String
    Dim lExeName As String
    Dim lChaveFull As String
    Dim lContX As Byte

    Err.Clear
    On Error GoTo TrataErro

    lChavesOffice = "SOFTWARE\Microsoft\Office\"

    lVersao(0) = "8.0\"
    lVersao(1) = "9.0\"
    lVersao(2) = "10.0\"
    lVersao(3) = "11.0\"
    lVersao(4) = "12.0\"
    lVersao(5) = "14.0\"

    Select Case pAplMS
    Case E_AplMS_Word
        lAplicativo = "Word\InstallRoot"
        lExeName = "winword.exe"
    Case E_AplMS_Excel
        lAplicativo = "Excel\InstallRoot"
        lExeName = "excel.exe"
    Case E_AplMS_PowerPoint
        lAplicativo = "PowerPoint\InstallRoot"
        lExeName = "powerpnt.exe"
    Case Else
        'Não previsto
        lFullPath = ""
        GoTo SAIDA
    End Select

    For lContX = LBound(lVersao) To UBound(lVersao)
        lChaveFull = lChavesOffice & _
                     lVersao(lContX) & _
                     lAplicativo
        lFullPath = gfRegistry_Read(lChaveFull, "Path", REG_EXPAND_SZ)
        lFullPath = Replace(lFullPath, """", "")
        If Len(lFullPath) > 0 Then
            If Right(lFullPath, 1) <> "\" Then lFullPath = lFullPath & "\"
            lFullPath = lFullPath & lExeName
            If Len(Dir(lFullPath)) > 0 Then
                'Aplicativo localizado!!!
                Exit For
            Else
                lFullPath = ""
            End If
        End If
    Next lContX

SAIDA:
    gfPathAplMicrosoft = lFullPath
    Exit Function

TrataErro:
    lMsgErr = "gfPathAplMicrosoft Erro: " & Err & "-" & Error
    lFullPath = ""
    GoTo SAIDA

End Function

