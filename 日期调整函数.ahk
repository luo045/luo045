; 日期时间调整函数
AdjustDateTime(InputDateTime, HoursToAdd, MinutesToAdd) {
    ; 补充分秒以确保正确解析
    DateTime := InputDateTime . "00"
    ; 添加小时
    if (HoursToAdd != 0)
        DateTime += %HoursToAdd%, Hours
    ; 添加分钟
    if (MinutesToAdd != 0)
        DateTime += %MinutesToAdd%, Minutes
    ; 格式化输出为yyyyMMddHHmm
    FormatTime, NewDateTime, %DateTime%, yyyyMMddHHmm
    return NewDateTime
}

; 示例使用
#1:: ; Win+1 测试延后3小时
    InputDateTime := "202310151230" ; 原始时间：2023-10-15 12:30
    NewTime := AdjustDateTime(InputDateTime, 3, 0)
    MsgBox 延后3小时后：%NewTime%
return

#2:: ; Win+2 测试提前半小时
    InputDateTime := "202310151230"
    NewTime := AdjustDateTime(InputDateTime, 0, -30)
    MsgBox 提前半小时后：%NewTime%
return

#3:: ; Win+3 自定义测试（示例：加2小时45分钟）
    InputDateTime := "202312250830" ; 2023-12-25 08:30
    NewTime := AdjustDateTime(InputDateTime, 2, 45)
    MsgBox 加2小时45分钟后：%NewTime%
return