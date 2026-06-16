Attribute VB_Name = "modHook"
Option Explicit

Private Declare Function SetWindowsHookEx Lib "user32.dll" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32.dll" (ByVal hHook As Long) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const CB_SETDROPPEDWIDTH    As Long = &H160

Private Type CWPSTRUCT
    lParam As Long
    wParam As Long
    message As Long
    hWnd As Long
End Type

Private Const WH_CALLWNDPROC As Long = 4&

Private m_X As Long
Private m_Y As Long
Private m_hHook As Long

Public Sub Hook(ByVal ThreadID As Long, ByVal X As Long, ByVal Y As Long)
    m_hHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf DlgPosProc, 0, ThreadID)
    m_X = X
    m_Y = Y
End Sub

Public Sub Unhook()
    Call UnhookWindowsHookEx(m_hHook)
End Sub

Public Sub DlgPosProc(ByVal n As Long, ByVal wParam As Long, ByRef Inf As CWPSTRUCT)
    Dim ClassName As String * 6
    If Inf.message = 1 Then
        Call GetClassName(Inf.hWnd, ClassName, Len(ClassName))
        If Left$(ClassName, 6) = "#32770" Then
            Call SetWindowPos(Inf.hWnd, 0&, m_X, m_Y, 0&, 0&, 1&)
        End If
    End If
End Sub

Public Sub SetDropDownWidth(mCombo As ComboBox)
    Dim RetVal As Long
    Dim PixelWidth As Long
    Dim MaxWidth As Long
    Dim LoopCounter As Long
    Dim lWidth As Long

    For LoopCounter = 0 To mCombo.ListCount - 1
        lWidth = mCombo.Parent.TextWidth(mCombo.List(LoopCounter))
        If lWidth > MaxWidth Then
            MaxWidth = lWidth
        End If
    Next LoopCounter

    MaxWidth = MaxWidth + (24 * Screen.TwipsPerPixelX)
'    If MaxWidth > (mCombo.Width * 2) Then
'        MaxWidth = (mCombo.Width * 2)
'    ElseIf MaxWidth < mCombo.Width Then
'        MaxWidth = mCombo.Width
'    End If

    PixelWidth = (MaxWidth \ Screen.TwipsPerPixelX)
    RetVal = SendMessage(mCombo.hwnd, CB_SETDROPPEDWIDTH, PixelWidth, 0)
    
End Sub
