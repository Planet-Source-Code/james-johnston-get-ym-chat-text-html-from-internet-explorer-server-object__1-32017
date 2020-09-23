Attribute VB_Name = "modEnumWindows"
Option Explicit

Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public YMText As String, YMHTML As String

Public Sub GetYMChatText()
    Dim lRet As Long, lParam As Long

    lRet = EnumWindows(AddressOf EnumWinProc, lParam)
End Sub

Function EnumWinProc(ByVal lhWnd As Long, ByVal lParam As Long) As Long
    Dim RetVal As Long
    Dim WinTitle As String
    
    WinTitle = GetWindowTitle(lhWnd)
    Debug.Print lhWnd, WinTitle
    
    If InStr(1, WinTitle, " -- Chat") Then
        YMText = GetIEText(lhWnd)
        YMHTML = GetIEHTML(lhWnd)
    End If
    
    EnumWinProc = True
End Function

Public Function GetWindowTitle(ByVal hWnd As Long) As String

    Dim l As Double
    Dim s As String
    l = GetWindowTextLength(hWnd)
    s = Space(l + 1)
    GetWindowText hWnd, s, l + 1
    GetWindowTitle = Left$(s, l)
    
End Function

