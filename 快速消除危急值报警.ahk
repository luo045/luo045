WinActivate, 瑞美专用浏览器 ahk_class WindowsForms10.Window.8.app.0.141b42a_r7_ad1
IfWinNotActive
{
    MsgBox, 无法激活瑞美专用浏览器窗口
    ExitApp
}

; 点击第一个按钮
Click, 475, 97
Sleep, 50 ; 减少等待时间

; 点击第二个按钮
Click, 1774, 57
Sleep, 50 ; 减少等待时间

; 点击确认按钮
Click, 1082, 604
Sleep, 100
Send,040545{Enter}{Enter}
Exit