Attribute VB_Name = "xxxWindowNT"
Option Explicit
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
                                    (ByVal lpBuffer As String, nsize As Long) As Long

Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" _
                                 (ByVal lpBuffer As String, nsize As Long) As Long
Public Function GetLoginComputerName() As String
    Dim cnt As Long
    Dim dl  As Long
    Dim S   As String

    cnt = 200
    S = String(cnt, 0)
    On Error GoTo 0
    dl = GetComputerName(S, cnt)

    GetLoginComputerName = Left(S, InStr(S, Chr(0)) - 1)
End Function
Public Function GetLoginUserName() As String
    Dim lngLEN  As Long
    Dim strUser As String

    strUser = Space(64)
    lngLEN = GetUserName(strUser, 64)
    GetLoginUserName = Left(Trim(strUser), Len(Trim(strUser)) - 1)
End Function
