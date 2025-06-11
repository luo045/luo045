#Persistent
CoordMode, ToolTip, Screen  ; 使用屏幕坐标系
提示内容 := "AHK助手运行中"   ; 自定义提示文字
刷新频率 := 50              ; 位置更新频率(毫秒)

SetTimer, 更新位置, %刷新频率%
return

更新位置:
    MouseGetPos, x, y
    x += 15  ; 水平右偏移
    y += 15  ; 垂直下偏移
    ToolTip, %提示内容%, x, y
return

; 退出热键
#q::ExitApp