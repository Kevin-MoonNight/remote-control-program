Option Explicit On
Module KeyboardAPI
    Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Integer, ByVal bScan As Integer, ByVal dwFlags As Integer, ByVal dwExtraInfo As Integer)
    Public Const keydown = &H100
    Public Const keyup = &H2
    Public Sub keyboard_down(ByVal hWnd As Integer)
        keybd_event(hWnd, 0, keydown, 0)
    End Sub
    Public Sub keyboard_up(ByVal hWnd As Integer)
        keybd_event(hWnd, 0, keyup, 0)
    End Sub
End Module
