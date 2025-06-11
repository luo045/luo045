SetKeyDelay, 100
#SingleInstance force
#Persistent
ExtractTime(searchText) {
	; 获取剪贴板内容
	clipContent := Clipboard
	pos := InStr(clipContent, searchText)
	; 查找关键文本位置
	if !(pos := InStr(clipContent, searchText))
		return ""

	; 提取并验证时间
	startPos := pos +7


	; 提取从 startPos 到 targetChar 左侧的字符
	extractedTime := SubStr(clipContent, startPos, 16)

	return extractedTime
}
time1:=ExtractTime("手术开始时间：")
time2:=ExtractTime("手术结束时间：")

; MsgBox,4, ,手术开始时间:%time1%+手术结束时间:%time2%
if RegExMatch(time1, "(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2})", match)
	yyyy := match1
mm := match2
dd := match3
hh := match4
min := match5

; MsgBox %yyyy%, %mm%, %dd%, %hh%, %min%
FillDateTimeField("|<>*165$56.070BVzy887y03A24220300k0V0V40kDzw8EyFbzs30244js301sDzxD10k0h08EG0DzwH8245js30MlUV0m20kAAAEECUUA030846zsS00k41222U", {x1:10, y1:200, x2:1060, y2:900}, yyyy, mm, dd, hh, min) ;

; 通用日期时间填写函数
; 通用日期时间填写函数

if RegExMatch(time2, "(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2})", match)
yyyy := match1
mm := match2
dd := match3
hh := match4
min := match5
; if RegExMatch(time2, "(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2})", match)
; yyyy := match1
; mm := match2
; dd := match3
; hh := match4
; min := match5


FillDateTimeField(Text:="|<>*163$56.070BUUE0k7y0349zzzw300k4V0300kDzzkEDz7zs309zWAE301s600X40k0h1vyDzDzwH8kUUS030MlU88/E0kAAAu2An0A030kzYAAS00k08830U", {x1:10, y1:200, x2:1060, y2:900}, yyyy, mm, dd, hh, min) ;
; 通用日期时间填写函数
FillDateTimeField(fieldText, fieldCoord, yyyy, mm, dd, hh, min, daysOffset := 0) {
	; 查找目标输入框
	if !(ok := FindText(X, Y, fieldCoord.x1, fieldCoord.y1, fieldCoord.x2, fieldCoord.y2, 0.1, 0.1, fieldText))
		return false

	; 点击输入框激活
	FindText().Click(X+520, Y+10, "L")
	Sleep, 100

	; 激活月份输入框
	if !(ok := FindText(XM, YM, 555-150000, 629-150000, 555+150000, 629+150000, 0, 0, "|<>*144$28.Dw000UE0021000Dw07wUE0DW100QDw00UUE0041000E40021k008"))
		return false

	FindText().Click(XM, YM, "L")
	Sleep, 500

	; 输入日期时间
	; Send, %mm% ;错误
	; ********************************这部分代替不能直接数字输入月份的办法

	if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
	{
		FindText().Click(X+498, Y+10, "L")
	}
	Sleep,100
	; 鼠标依次点击年,月,日,时间
	; 点击月份输入框箭头,激活月份输入框
	Text:="|<>*144$28.Dw000UE0021000Dw07wUE0DW100QDw00UUE0041000E40021k008"
	if (ok:=FindText(X, Y, 555-150000, 629-150000, 555+150000, 629+150000, 0, 0, Text))
	{
		FindText().Click(X, Y, "L")
	}
	Sleep,300
	; Send, %mm% ;错误

	Send, {TAB}
	Send, %yyyy%
	Send, {TAB}
	Send, %hh%
	Send, {TAB}
	Send, %min%
	Send, {Esc}
	Send, {TAB}
	; ********************************这部分代替不能直接数字输入月份的办法
	Currentmonth :=A_Mon+0
	Targetmonth :=mm
	EnvSub,Targetmonth,%Currentmonth%
	Send,{Left}
	if (Targetmonth != 0) {
		direction := (Targetmonth > 0) ?"Right" : "Left"
		loops := Abs(Targetmonth)
		Send, {%direction% %loops%}
	} else {
		Sleep, 300
	}
	; Send, {Enter}
	Send, {Esc}



	; ***********************输入日期
	Text:="|<>*204$10.Fz44EFt4I1F7Xc" ;点击15号
	if (ok:=FindText(X, Y, 0, 0, 0, 0, 0, 0, Text))
	{
		FindText().Click(X, Y, "L")
	}
	Sleep,300
	; 设置基准日期（当天）
	CurrentDate :=15
	TargetDate :=dd
	EnvSub,TargetDate,%CurrentDate%
	if (TargetDate != 0)
	{
		direction := (TargetDate > 0) ? "Right" : "Left"
		loops := Abs(TargetDate)
		Send, {%direction% %loops%}
	}
	else
		Sleep,300
	Send, {TAB}
	; ***************************************输入日期

	Send {Enter}        ; 当天直接确认
	return true
}
ExitApp