Option Explicit On
Module MouseAPI
    'API定義
    Private Declare Sub mouse_event Lib "user32" Alias "mouse_event" (ByVal dwFlags As Integer, ByVal dx As Integer, ByVal dy As Integer, ByVal cButtons As Integer, ByVal dwExtraInfo As Integer)
    Private Declare Function SetCursorPos Lib "user32" (ByVal X As Integer, ByVal Y As Integer) As Integer
    Public Const mouse_left_down = &H2
    Public Const mouse_left_up = &H4
    Public Const mouse_middle_down = &H20
    Public Const mouse_middle_up = &H40
    Public Const mouse_right_down = &H8
    Public Const mouse_right_up = &H10
    Public Const mouse_move = &H1
    Public Const mouse_wheel = &H800
    Structure PointAPI
        Dim X, Y As Integer
    End Structure
    Public Sub move_mouse(ByVal xMove As Integer, ByVal yMove As Integer) '移動滑鼠
        setcursorpos(xMove, yMove)
    End Sub
    Public Sub left_down() '按下滑鼠左鍵
        mouse_event(mouse_left_down, 0, 0, 0, 0)
    End Sub
    Public Sub left_up() '放開滑鼠左鍵
        mouse_event(mouse_left_up, 0, 0, 0, 0)
    End Sub
    Public Sub middle_down() '按下滑鼠中鍵
        mouse_event(mouse_middle_down, 0, 0, 0, 0)
    End Sub
    Public Sub middle_up() '放開滑鼠中鍵
        mouse_event(mouse_middle_up, 0, 0, 0, 0)
    End Sub
    Public Sub right_down() '按下滑鼠右鍵
        mouse_event(Mouse_Right_Down, 0, 0, 0, 0)
    End Sub
    Public Sub right_up() '放開滑鼠右鍵
        mouse_event(mouse_right_up, 0, 0, 0, 0)
    End Sub
    Public Sub wheel(ByVal ScrollValue As Integer) '滾輪
        mouse_event(mouse_wheel, 0, 0, ScrollValue, 0)
    End Sub
End Module
