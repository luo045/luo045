#Persistent
CoordMode, Pixel, Screen  ; 设置坐标模式为屏幕绝对坐标

; 定义滚轮向下翻页热键
WheelDown::
    Loop {
        ; 使用FindText函数检测题目位置（需提前截取题目特征图）
        if (FindText(X, Y, 0, 0, A_ScreenWidth, A_ScreenHeight, 0.9, 0.9, "题目特征图路径.png")) {
            ; 判断是否到达目标区域（例如屏幕中央）
            if (Y > 300 && Y < 500) {  ; 假设目标区域Y轴为300-500像素
                MsgBox 题目已到达目标位置！
                break
            }
        }
        ; 未找到或未到位时继续滚动
        Send {WheelDown 3}  ; 每次滚动3格
        Sleep 500  ; 等待页面稳定
    }
return