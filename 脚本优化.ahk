#SingleInstance Force
SetKeyDelay, 200
SetWorkingDir %A_ScriptDir%
CoordMode, Mouse, Screen

; ---------------------------
; 配置区
; ---------------------------
global CONFIG := {
    searchRange: 500,     ; 合理缩小查找范围
    clickOffsetX: 498,     ; 水平点击偏移
    retryCount: 3,        ; 操作重试次数
    screenWidth: 1920,     ; 屏幕分辨率适配
    screenHeight: 1080
}

; 文本模板库
global TEXT_PATTERNS := {
    mainDiagnosis: "|<>*170$128.EA4kBUwT00A73zw8kTznzm71hwMtYQ0738aH460F0k4H8Lc6M11U3kX9YlTw4EA11ZYu1Y0E80AMGNAk094nzyWByUN0420364aHAUWF8k4X8HDqE10U0lV94r88oGA19YZmFY0E9wAMGFAn455Xzm6NyYN0420364gTAF1FEk4W8PN6M11U0lXC0n4kQQA1+A4mFX6Fk0AAW0Al81430HC1z4MT7k031kU3A20F0zw400F00000000DznzzzzA100000000000020A0000008",
    surgeryStart: "|<>*163$51.0U7ztUwETzy48AQm3/50V1W0EN7848AU23/50V1Y0ENztzzgU230U0V1Y0EPzw48AU23F8UV1a0EGTY88AMnaE0W11Vw7W0MU80000U",
    icuEntry: "|<>*174$51.300000000808EA7W30U111XaEM60Pzgk230k311Y0EMD0s8AU2318D11Y0EMNUPzAU2324311Y0EMkkM8Ak32A3311X6Qn0AM8ADVwk0Xzw0004"
}

; ---------------------------
; 核心函数库
; ---------------------------
; 带重试的智能点击
ClickByText(textPattern, xOffset:=0, yOffset:=0, retries:=CONFIG.retryCount) {
    loop % retries {
        if (ok := FindText(X, Y, 
            0, 0, CONFIG.screenWidth, CONFIG.screenHeight, 
            0, 0, textPattern))
        {
            FindText().Click(X + CONFIG.clickOffsetX + xOffset, Y + yOffset, "L")
            return true
        }
        Sleep 500
    }
    MsgBox 元素未找到：%textPattern%
    return false
}

; 安全发送输入
SafeSend(keys, delay:=100, retries:=3) {
    try {
        SetKeyDelay %delay%
        loop % retries {
            Send % keys
            if !ErrorLevel
                return true
            Sleep 300
        }
    }
    MsgBox 输入失败：%keys%
    return false
}

; 获取手术时间
GetSurgeryTime() {
    if !ClickByText(TEXT_PATTERNS.surgeryStart, 0, 4) 
        return

    SafeSend("{Esc}^a^c")
    ClipWait 2
    if ErrorLevel {
        MsgBox 剪贴板获取超时
        return
    }
    
    if RegExMatch(Clipboard, "(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2})", match) {
        return {
            year: match1,
            month: match2,
            day: match3,
            hour: match4,
            minute: match5
        }
    }
    MsgBox 时间格式解析失败
    ExitApp
}

; ---------------------------
; 主流程
; ---------------------------
#Persistent
SetTitleMatchMode, 2

; 输入体外循环时间
InputBox, cpbhh1, 体外循环开始 hh
InputBox, CPBmin1, 体外循环开始 mm
InputBox, cpbhh2, 体外循环结束 hh
InputBox, CPBmin2, 体外循环结束 mm

; 诊断信息填写
ClickByText(TEXT_PATTERNS.mainDiagnosis)
SafeSend("{DOWN 2}{Enter}")

; 获取手术时间
surgeryTime := GetSurgeryTime()

; ICU时间填写
if !FillDateTimeField(TEXT_PATTERNS.icuEntry, surgeryTime, 30) {
    MsgBox ICU时间填写失败
    ExitApp
}

; 其他业务流程...
return

; ---------------------------
; 日期时间处理模块
; ---------------------------
FillDateTimeField(fieldPattern, baseTime, addMinutes:=0) {
    adjustedTime := AdjustDateTime(baseTime, addMinutes)
    
    if !ClickByText(fieldPattern) 
        return false

    ; 日期选择器交互逻辑
    if !ClickCalendarMonth(adjustedTime.month)
        return false

    SafeSend(adjustedTime.day)
    SafeSend("{Tab}" adjustedTime.hour "{Tab}" adjustedTime.minute)
    return true
}

AdjustDateTime(baseTime, addMinutes) {
    formattedTime := baseTime.year baseTime.month baseTime.day 
        . baseTime.hour baseTime.minute
    formattedTime += addMinutes, Minutes
    
    FormatTime, newTime, %formattedTime%, yyyyMMddHHmm
    RegExMatch(newTime, "(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})", match)
    
    return {
        year: match1,
        month: match2,
        day: match3,
        hour: match4,
        minute: match5
    }
}

ClickCalendarMonth(targetMonth) {
    if !ClickByText("|<>*144$28.Dw000UE0021000Dw07wUE0DW100QDw00UUE0041000E40021k008")
        return false
    
    currentMonth := 3  ; 假设当前显示3月
    delta := targetMonth - currentMonth
    if (delta != 0) {
        dir := (delta > 0) ? "Right" : "Left"
        SafeSend("{" dir " " Abs(delta) "}")
    }
    return true
}

; ---------------------------
; 异常处理
; ---------------------------
OnError(exception) {
    MsgBox 发生错误：%exception%
    ExitApp
}