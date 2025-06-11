^!v::  ; 热键：Ctrl+Alt+V触发解析
	Clipboard = %Clipboard%  ; 获取剪贴板内容
	Clipboard := Trim(Clipboard)  ; 去除首尾空格
	; 正则表达式解析（兼容V1语法）
	if RegExMatch(Clipboard, "(\d{4})-(\d{1,2})-(\d{1,2}) (\d{1,2}):(\d{2})", match) 
		yyyy := match1
		mm := Format("{:02}", match2)  ; 自动补零
		dd := Format("{:02}", match3)
		hh := Format("{:02}", match4)
		min := match5
		InputDateTime=%yyyy%%mm%%dd%%hh%%min%00
	return

; 日期时间调整函数
AdjustDateTime(InputDateTime, HoursToAdd, MinutesToAdd) {
	if (HoursToAdd != 0)
		InputDateTime += %HoursToAdd%, Hours
	; 添加分钟
	if (MinutesToAdd != 0)
		InputDateTime += %MinutesToAdd%, Minutes
	; 格式化输出为yyyyMMddHHmm
	FormatTime, NewDateTime, %InputDateTime%, yyyyMMddHHmm
	return NewDateTime
}

#3:: ; Win+3 自定义测试（示例：加2小时45分钟）
	; InputDateTime := "202312250830" ; 2023-12-25 08:30
	NewTime := AdjustDateTime(InputDateTime, 2, 45)
	 MsgBox 加2小时45分钟后：%NewTime%
return





