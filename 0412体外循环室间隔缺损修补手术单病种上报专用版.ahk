#SingleInstance Force
SetKeyDelay, 200

#SingleInstance force
#Persistent
SetWorkingDir %A_ScriptDir%
SetTitleMatchMode, 2
CoordMode, Mouse, Screen


MsgBox,4096,输入体外循环时间 格式为
InputBox,cpbhh1,体外循环开始 hh
InputBox,CPBmin1,体外循环开始 mm
InputBox,cpbhh2,体外循环结束 hh
InputBox,CPBmin2,体外循环结束 mm
MsgBox,4096,手动返回输入框界面 再点击继续
; ********************************
; 查找 主要诊断ICD-10四位亚目编码与名称：
Text:="|<>*170$128.EA4kBUwT00A73zw8kTznzm71hwMtYQ0738aH460F0k4H8Lc6M11U3kX9YlTw4EA11ZYu1Y0E80AMGNAk094nzyWByUN0420364aHAUWF8k4X8HDqE10U0lV94r88oGA19YZmFY0E9wAMGFAn455Xzm6NyYN0420364gTAF1FEk4W8PN6M11U0lXC0n4kQQA1+A4mFX6Fk0AAW0Al81430HC1z4MT7k031kU3A20F0zw400F00000000DznzzzzA100000000000020A0000008"
;诊断ICD-10四位亚目

if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
	FindText().Click(X+498, Y, "L")
}
Sleep,200
; 鼠标依次诊断ICD-10四位亚目输入框
Send, {DOWN}{DOWN}{Enter} ;选择第2项 房间隔缺损
; ********************************


; ********************************
; 查找 诊断ICD-10六位临床扩展编码 与名称：
Text:="|<>*169$187.862M6kSDU063U404M140208A7zkUUzz271hwMtYQ073810466a01U42208bz40UaEjEAk2307V6002zvHyTzjztzwG0W8E6KHc6E10U0lV00301eE80180U2Lzl4FoFjo380UE0MkjztY4q4440Y0F4CU1288m4nxY0E80AME01m2O03zwO0DzWTwxz4mGt8m084y6A88UNWBTt3Ut04F3Aem0m6NyYN0420364A8AF6hYWc4U3zxyJN0F4EqmAk2301X6A669XKmGG2E1/43zYzsckH96AN700km41X4VfN+8V80YwDZGEANkDsX3sy00MC40NUEpzi48c0Xg9Gds61004E0000000404zz2kI21o0F0kcQYSU"
;诊断ICD-10六位临床扩展编码
if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
	FindText().Click(X+498, Y, "L")
}
Sleep,100
Send, {Enter} ;默认 房缺
; ********************************


; ********************************
; 查找 主要手术操作栏中提取ICD-9-CM-3四位亚目编码与名称：
Text:="|<>*166$251.9z140l40k2Dvy0MD7k0700yM303Vzy487zszw88DzkH22E1W81U4En/wln8M0H034s609W94881Y108Lz20Xzw8zzUVzyyzaK9a0EM1W0A1kQ014G8jy3820FA28F100GE6Tu64F3DYH80UE260E3Uc028YF00aF7zYzwEW2RtgUA04A8vwN9aE10U4A1U6XE0AF8aEFga81Dk0XADeLFwQ08MHU0yGAU210AM30B4U3kWFQUVN8E25znrwtrW21s0EkizxYoN042TDrq0P9DUl8W8V2mEzwHeaU8EE447bwzz4238km084010A0mW00mlwF47Z109zJR0Ejz8Dn01328bbxVa0EM020A1b401b08X87+20EDyPzd3UEE60060H8QbX6AVU0A0QH680a80F2k6E40XpIo12+EUUA00A0jk1Na7lw03U0D68E1sE0W10AUDzBeds2R4N10PzUM7Fy3V00000000000000zz5zzzyE22FnEs000000000000000000000000000102000000000001"

;主要手术操作栏中提取ICD-9-CM-3四位亚目编码与名称：
if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
	FindText().Click(X+498, Y, "L")
}
Sleep,100
Send, {DOWN}{Enter} ;选第一项 切开法
; ********************************

; ********************************
; 查找主要手术操作栏中提取ICD-9-CM-3六位临床扩展编码与名称：
Text:="|<>*167$279.9z140l40k2Dvy0MD7k0700yM303U204814030847zsUUzz1A89068U60F3Ajn7AVU1A0AHUM0a0E0UUcU0810Uk35zkU8zz2Dzs8TzjjtZWNU460MU30Q700E10/zZDxzyyzrztA28F100GE6Tu64F3DYH80UE260E3Uc0200100fEA0140k3Hzl289ram0k0EkXjlYaN0420Ek60OB00nzzN15F1VU+U6F3w08n3uZoT70264s0DYX80UE360k3F80w00788c0DzlY0zz5znrwtrW21s0EkizxYoN042TDrq0P9DUk8E8V5DtVUMU6F1CeO0V10EESTnzwE8AX380UE040k3+8033214Ed9AC5A0zzTpLE4/zm3wk0EkW9tzMNU4600U30Nl00ME88m599Zc904YEDyPzd3UEE60060H8QbX6AVU0A0QH680a40V2kd9/Al80YwDJHE48d220k00k2z05aMT7k0C00wMV07V06845DtFW/04xXOeS0bF6EE6zs61oTUsE000000000000080FTy91EA7E1632FnEwU"

;主要手术操作栏中提取ICD-9-CM-3六位临床扩展编码与名称：
if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
	FindText().Click(X+498, Y, "L")
}
Sleep,100
Send, {DOWN}{DOWN}{Enter} ;选第2项房间隔缺损人造补片修补术 胸腔镜为第5项



; ********************************
; 是否为出院后31天内非预期重复住院：单项选择是否/男女：
Text:="|<>*154$114.Dz7zsW00U7W00SD00zy810A0G08V400zk11U10Dz0R0208V5TsU01XU10813Yk208V5E8U01UU10Dz449zyDz6Dszz10Vzz00000220V500U0C0U10Tznzk22EUZzsU01UU200E20E4WEUY4Ujw0UU2U4TW0EAGEUZ4Uc40UU4EAE20EM2EUb4Z84FUU88SE3zkk6TzY8ZDwD0Uk6lzm0F0w00Yku8400103U"
;是否为出院后31天(问题的前8个字)：
if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
	;FindText().Click(514, Y, "L") ;点击 是
	FindText().Click(600, Y, "L") ;点击 否
}
Sleep,100

; ********************************
Sleep,300
WinWait, 郑大一单病种质控上报系统 - Google Chrome, Chrome Legacy Window
IfWinNotActive, 郑大一单病种质控上报系统 - Google Chrome, Chrome Legacy Window, WinActivate, 郑大一单病种质控上报系统 - Google Chrome, Chrome Legacy Window
WinWaitActive, 郑大一单病种质控上报系统 - Google Chrome, Chrome Legacy Window
MouseClick, left,  37,  387
Sleep, 100

Send, {PGDN}
Sleep,300
; MsgBox,4096,,确认手术时间正确 再继续
; MsgBox,4096,确认手术时间正确 再继续
getdata(){

	Sleep, 100
	; MouseClick, left,  291,  299
	; Sleep, 100

	; Send, {PGDN}
	Sleep,200
	; 查找 手术结束 复制并发送剪贴板
	Text:="|<>*163$56.070BUUE0k7y0349zzzw300k4V0300kDzzkEDz7zs309zWAE301s600X40k0h1vyDzDzwH8kUUS030MlU88/E0kAAAu2An0A030kzYAAS00k08830U"


	if (ok:=FindText(X, Y, 153-150000, 641-150000, 153+150000, 641+150000, 0, 0, Text))
	{
		FindText().Click(X+520, Y+4, "L")
	}
	Sleep, 100
	Send, {Esc}{CtrlDown}a{CtrlUp}{CtrlDown}c{CtrlUp}
	ClipWait
	Clipboard := Trim(Clipboard)  ; 去除首尾空格
	; 正则表达式解析（兼容V1语法）
	if RegExMatch(Clipboard, "(\d{4})-(\d{1,2})-(\d{1,2}) (\d{1,2}):(\d{2})", match)
		yyyy := match1
	mm := match2
	dd := match3
	hh := match4
	mms := match5
	InputDateTime=%yyyy%%mm%%dd%%hh%%mms%00 ;把手术结束时间发送到inputdatetime
	; 示例时间参数 - 把手术结束时间
	NewTime := AdjustDateTime(InputDateTime, 0, 30)
	RegExMatch(NewTime, "(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})", match) ;把手术结束时间延后30min发送到newtime，yyyyMMddhhmm格式 match1-5表示手术结束时间：年，月，日，时，分 延时30分钟
	; global yyyy, mm, dd, hh, min  ; 添加全局声明
	yyyy := match1
	mm := match2
	dd := match3
	hh := match4
	min := match5
	; MsgBox,%yyyy%%mm%%dd%%hh%%min% 手术结束半小时的时间
	return {yyyy: yyyy, mm: mm, dd: dd, hh: hh, min: min}
}

getdata1(){

	Sleep, 100
	; MouseClick, left,  291,  299
	; Sleep, 100

	; Send, {PGDN}
	Sleep,200
	; 查找 手术开始  日期时间： 复制并发送剪贴板
	Text:="|<>*165$56.070BVzy887y03A24220300k0V0V40kDzw8EyFbzs30244js301sDzxD10k0h08EG0DzwH8245js30MlUV0m20kAAAEECUUA030846zsS00k41222U"


	if (ok:=FindText(X, Y, 153-150000, 641-150000, 153+150000, 641+150000, 0, 0, Text))
	{
		FindText().Click(X+520, Y+4, "L")
	}
	Sleep, 300
	Send, {Esc}{CtrlDown}a{CtrlUp}{CtrlDown}c{CtrlUp}
	ClipWait
	Clipboard := Trim(Clipboard)  ; 去除首尾空格
	; 正则表达式解析（兼容V1语法）
	if RegExMatch(Clipboard, "(\d{4})-(\d{1,2})-(\d{1,2}) (\d{1,2}):(\d{2})", match)
		yyyy := match1
	mm := match2
	dd := match3
	hh := match4
	mms := match5
	InputDateTime=%yyyy%%mm%%dd%%hh%%mms%00 ;把手术结束时间发送到inputdatetime
	; 示例时间参数 - 把手术结束时间
	NewTime := AdjustDateTime(InputDateTime, 0, -20)
	RegExMatch(NewTime, "(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})", match) ;把手术结束时间提前20min发送到newtime，yyyyMMddhhmm格式 match1-5表示手术结束时间：年，月，日，时，分 延时30分钟
	; global yyyy, mm, dd, hh, min  ; 添加全局声明
	yyyy := match1
	mm := match2
	dd := match3
	hh := match4
	min := match5
	; MsgBox,%yyyy%%mm%%dd%%hh%%min% 手术开始的时间
	return {yyyy: yyyy, mm: mm, dd: dd, hh: hh, min: min}
}
;得到手术开始时间
getdata2(){



	time3:=ExtractTime("体外循环开始")
	time4:=ExtractTime("体外循环结束")

	MsgBox 体外循环开始时间：%time3%,体外循环结束时间%time4%
	if RegExMatch(time3, "(\d{2}):(\d{2})", match)
	; yyyy := match1
	; mm := match2
	; dd := match3
		CPBhh1 := match1
	CPBmin1 := match2
	MsgBox 开始CPB %CPBhh1%  %CPBhh1%


	if RegExMatch(time4, "(\d{2}):(\d{2})", match)
	; yyyy := match1
	; mm := match2
	; dd := match3
		CPBhh2 := match1
	CPBmin2 := match2
	MsgBox 终止CPB %CPBhh2%  %CPBmin2%
	return {CPBhh1 : hh, CPBmin1: mm, CPBhh2 : hh, CPBmin2: mm}
}

; 时间调整函数
AdjustDateTime(InputDateTime, HoursToAdd, MinutesToAdd) {
	if (HoursToAdd != 0)
		InputDateTime += %HoursToAdd%, Hours
	if (MinutesToAdd != 0)
		InputDateTime += %MinutesToAdd%, Minutes
	FormatTime, NewDateTime, %InputDateTime%, yyyyMMddHHmm
	return NewDateTime
}

; 通用日期时间填写函数
FillDateTimeField(fieldText, fieldCoord, yyyy, mm, dd, hh, min, daysOffset := 0) {
	; 查找目标输入框
	if !(ok := FindText(X, Y, fieldCoord.x1, fieldCoord.y1, fieldCoord.x2, fieldCoord.y2, 0, 0, fieldText))
		return false

	; 点击输入框激活
	FindText().Click(X+520, Y+10, "L")
	Sleep, 100

	; 激活月份输入框
	if !(ok := FindText(XM, YM, 555-150000, 629-150000, 555+150000, 629+150000, 0, 0, "|<>*144$28.Dw000UE0021000Dw07wUE0DW100QDw00UUE0041000E40021k008"))
		return false

	FindText().Click(XM, YM, "L")
	Sleep, 100

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
		loops := Abs(TargetDate)+daysOffset
		Send, {%direction% %loops%}
	}
	else
		Sleep,300
	Send, {TAB}{Esc}
	; ***************************************输入日期

	Send {Enter}{Esc}        ; 当天直接确认
	return true
}
ExtractTime(searchText) {
	; 获取剪贴板内容
	clipContent := Clipboard
	pos := InStr(clipContent, searchText)
	; 查找关键文本位置
	if !(pos := InStr(clipContent, searchText))
		return ""

	; 提取并验证时间
	startPos := pos -5


	; 提取从 startPos 到 targetChar 左侧的字符
	extractedTime := SubStr(clipContent, startPos, 5)

	return extractedTime
}

data := getdata()
Send,{Esc}
Sleep,300
Send,{Esc}
data1:= getdata1()
; data2:= getdata2()
Send,{Esc}
FillDateTimeField("|<>*174$51.300000000808EA7W30U111XaEM60Pzgk230k311Y0EMD0s8AU2318D11Y0EMNUPzAU2324311Y0EMkkM8Ak32A3311X6Qn0AM8ADVwk0Xzw0004", {x1:10, y1:200, x2:1060, y2:900}, data.yyyy, data.mm, data.dd, data.hh, data.min) ;入住ICU：
Send {Esc}
FillDateTimeField("|<>*163$51.0U7ztUwETzy48AQm3/50V1W0EN7848AU23/50V1Y0ENztzzgU230U0V1Y0EPzw48AU23F8UV1a0EGTY88AMnaE0W11Vw7W0MU80000U", {x1:10, y1:200, x2:1060, y2:900}, data.yyyy, data.mm, data.dd+1, 10, data.min) ;离开ICU
Sleep, 100
WinWait, 郑大一单病种质控上报系统 - Google Chrome, Chrome Legacy Window
IfWinNotActive, 郑大一单病种质控上报系统 - Google Chrome, Chrome Legacy Window, WinActivate, 郑大一单病种质控上报系统 - Google Chrome, Chrome Legacy Window
WinWaitActive, 郑大一单病种质控上报系统 - Google Chrome, Chrome Legacy Window
MouseClick, left,  37,  387
Sleep, 100

Sleep, 100
Send, {PGDN}
Sleep,300
; ASD-1：术前病情严重程度评估
; 1********************************
; 是否实施手术前的超声心动图评估：：单项选择
Text:="|<>*154$214.Dz7zs208E01k0E20120Fy0E0800ETzYTsEU0U40k040F0Ts00E8E88129zz0E7t142801203zk7ETzbzw0000DzxwyT8U4010044zcY2/zk813Yl0290001zy004I8EWDzUWE1yH4UEEUU4zwEEVW0UETz0A1yGFUV5k00+1019GWsU620E0000183VM001s4994Wz03zl86TYYC8U0s80TznzkkU+QU008ETYbl81y8F4U8EGK7Wzwby010810Y0DGzzl0V2GF0Y48zwG0W3962802E817sU4TzoVk00M1bt942HkW0E8198Y68U090UAE20E30G400102EYYE91y800U1AWH0WU0Y25t0Dz0nV910U00122T0bY900226A91mA02TsQTwU4M0fbsS0004tl4wrzk00Ds0HbzsU090U0000000000000000000000000020E0U00002"

;是否实施手术前的超声心动图评估：：
if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
	; FindText().Click(514, Y, "L") ;点击 是
	FindText().Click(600, Y, "L") ;点击 否
}
Sleep,100
; ********************************

; 2********************************
; 是否实施手术前的心导管检查评估：
Text:="|<>*166$115.00000000000000007zUTzV3E0ESM04020GY30E80EV4z8911zy0U1W1zs00FEU044j2U0zzzvw003w8zzUDvkYzzUY0l4zzW1zAA01B8KMA0GVwWAl102263wabi4639BZu7zUTt330EGGK6zt02A53AE4YXXEEND/ZFYTwTnVzs2ELF898YdEdu244EU3019q9W/YGQYL70w1sszz2Y25VaC9AHG0Uz0wq0k1Xz3UE6RMs13nUzUVzzs00000200000000000008"
;是否实施手术前的心导管检查评估：
if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
	; FindText().Click(514, Y, "L") ;点击 是
	FindText().Click(600, Y, "L") ;点击 否
}
Sleep,100
; ********************************
; ********************************


; ASD-2：手术适应证 手术方法的选择：
Text:="|<>*165$100.0Q0q0A10k484/0by3y03A0E630UU9A28M0k0A3zzBzbrsbyyH030zzkU00kFEUn0Us7zs302063162sA2DkEk0S0Dw8A4G8jzDVk302o0UkDzTAW9XUE3zwH8232214G8aGTs0k6AMMA8G4E8aO84030kkl0l24F0Wksjz4A030A24Ftw2Q0210LU0A1XsHsYHkDzs400000A000000000002"

;手术方法的选择：
if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
	FindText().Click(X+377, Y, "L")
}
Send, {UP}{Enter} ;默认是封堵术，选择直视下房间隔缺损修补术
; ********************************


; 常用房间隔缺损封堵器的材料：
Text:="|<>*159$185.827zk20HzDzt0UFz20V10yT240U8820888EXzsE2I0610W2412294W8810GJ43zwEV40G04dyDW7rwz24TXtwyzDzul8408zyDzYz9G0ITW88Fzy+01152A35UEPzF24E192H7s89400U8HzUF2A4Q628UY2248UUHwZ03yGDzjuEUUTzoG8oIzN0Dw7zlzyY9+zkUZm124V3w2kDaGUcMG0008EW818GJGZTwYG4B28/kDF0h2EsD1zwEV8zWTYuJ+W18Yz+7TntwW128Wjs20912F14U97yJ+2H8E4kUYG9424V10EI3W24a280G8Yzm4BUw81zDbns4+220UU005uEwE7YFs33vUi3k22F8YHkEs410U"


;手术方法的选择：
if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
	FindText().Click(X+377, Y, "L")
}
Send, {down}{down}{down}{down}{down}{down}{Enter} ;选择无法确定，或者无记录
; ********************************


; 4********************************
; ASD-3：术中验证房间隔缺损手术效果的措施
; 术中是否使用经食道超声：
Text:="|<>*153$157.08081zszz241zw800M04E8z08001040U40k1zwV28z0O1Dz48bzw003zwTy0u110EV41UIUEEDYE207zt12813YkbyDzYV0lA1zV28zy460UV7zW24mF48G3tzxQUEWs0027UEEU0000N8W4/a3422Tvw0Dz04888HzyTy4zlzw003z1840z48U427zs1081228V2SzV0UbyEEXzkA0m248z40VB0EVM20zkG19sF0840810AE20EVUEEU10KA8z47sU04000UD81zsFw88FkU8sC03W4U02000EATwU4/3U5v7y5b0zzTz00000000000000000002000000000E"
;术中是否使用经食道超声：
if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
	; FindText().Click(514, Y, "L") ;点击 是
	FindText().Click(600, Y, "L") ;点击 否
}
; ********************************
Sleep,200
; 5********************************
; 是否使用体外循环：
Text:="|<>*157$116.Dz7zsEUDzV20EE43bjs20E307zm48EU442S0UE0zw1o220V2/zXt1YE8A0813YkbyDzW70WE5zm20Hzl12N8W49VkEa2F1tU400006G8V2M+59MZy8S0TznzkbyDzW+UQHN0WBE040U488W48YY342LsaI0Fy812O0V2990V0c2D4UAE20EVUEEWjcEE+0a10LY0zw8y448kV842bu0E77z812ks1S88410U204000000000022000000008"

;是否使用体外循环：
if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
	FindText().Click(514, Y, "L") ;点击 是
	; FindText().Click(600, Y, "L") ;点击 否
}
Sleep,600

Send,{Esc}
WinWait, 郑大一单病种质控上报系统 - Google Chrome, Chrome Legacy Window
IfWinNotActive, 郑大一单病种质控上报系统 - Google Chrome, Chrome Legacy Window, WinActivate, 郑大一单病种质控上报系统 - Google Chrome, Chrome Legacy Window
WinWaitActive, 郑大一单病种质控上报系统 - Google Chrome, Chrome Legacy Window
MouseClick, left,  37,  387
Sleep, 100

Send, {PGDN}
Sleep,500


MsgBox,4096,点击继续填写体外循环
Sleep,600
FillDateTimeField("|<>*163$88.4TW20zwNjU0Eby0E28820HzWT11087s8V4816O9A4E0U40jYMU4TjYzxDm4FyGzW0FaWH14V8Ty1D1Dz7u9x4HwU6M4U0U4NcYoF8G1NXKzW0FayH94V85y8m281Dy9A4HwUSDXc8U498bkF825s0OzXzlb2H1408Rzy228104s0sE7UU", {x1:10, y1:200, x2:1060, y2:900},	data.yyyy, data.mm, data.dd, CPBhh1, CPBmin1)
Send,{Esc}
FillDateTimeField("|<>*164$88.AE040zwNjU0Eby0Vy0E20HzWT11084AMF0816O9A4E0ULP140U4TjYzxDm7ok4TW0FaWH14V8G7kF0Dz7u9x4HwUFUl40U4NcYoF8G3lk4E20FayH94V800kF081Dy9A4HwUCk140U498bkF8271s4E3zlb2H1408E0Hzz8144s0sE7UU", {x1:10, y1:200, x2:1060, y2:900},	data.yyyy, data.mm, data.dd, CPBhh2, CPBmin2)

Send,{Esc}
; ********************************

; MsgBox
; 6********************************
; 术中是否再次阻断主动脉：
Text:="|<>*164$158.0q0A0zwTzbzs20SzYkA20047a00Ak30A10A030EU4cNhw0k7t188030TzXzkDkDz2TtG6Lc0400EGw0zzoA8k4Cn2AE42IVYm1zy0TrV8EA132Dz4AAX409bDtyU0k01B8K47UEkU0000Dz12FG6HDkA1yGHr02o4A/zzDz2AEkUKVZmE3044YZU1AVzy0E20EX48I4jtiYTzW39tQ1X6EkV7sU4zzoAV+6PN0A0YWGZ0kkkA0l08120F28QVYmE30L8Yt8EA030SE3zkU431Y8Nz40k6C9AH4300kATwU48D109Dz0FDzw3CgQ00000000000000000000000U0008"
;术中是否再次阻断主动脉：
if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
	; FindText().Click(514, Y, "L") ;点击 是
	FindText().Click(600, Y, "L") ;点击 否
}

; ********************************

; 6********************************
; 是否围术期使用血制品：
Text:="|<>*154$144.Dz7ztzy0423sEUDzU80c2Dz0810A122017y8Tz8EUE0cG810Dz0R1Tu00228UU8EXzlzG810813Yl22TzXnsbyDzWGF8GDz1Dz449Tm0k229YW8EWGE8G81100001221s3m9YW8EWGHzm000Tznzlzu24228byDzWGE8GTDU0E20F2+4223sUW8EWGFzGF8U4TW0F2+M1jy8aU8EWGF9GF8UAE20F2mE0WG8VUEEWGF92F8VSE3zl2200448XsEEWGFD2TDVlzm0Fzy0004sgC0Ljzw8CF8U000010200000000000000000U"
;是否围术期使用血制品：
if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
	; FindText().Click(514, Y, "L") ;点击 是
	FindText().Click(600, Y, "L") ;点击 否
}
; send,43
; ********************************


;术后是否实施机械通气：
Text:="|<>*163$144.0q01szwTzU80V0Ay21F7y4000l3z0k40k040H0AW218WA7zU0k200zw1w7zvzzyWDzw1k800zzm00k4Cn408Y0AW21U7yTz10k3zwzwEklW0UEAW3JfYWE011s20000000G0vKSW6ocbyTy02o203zzDz321fmSW6IsYW0204m2zs10810a1jGwW+zkby020MlaUMFy81Dzx/QAWGIUYW02EkkoUMl0810A1/EAWGJYYy01F0k4ztt0Dz0nV/1B2GLRk001l0k8UP7z8170OtyC1WY8Dz01UU"

if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
	FindText().Click(514, Y, "L") ;点击 是
	; FindText().Click(600, Y, "L") ;点击 否
}
; ********************************
MsgBox,4096,点击继续机械通气时间 注意不要和体外循环时间在同一屏幕内

; FillDateTimeField("|<>*164$60.4TW20049zU40W21w440UTUW4FA4E0U40jYNDzHwV4TYjtA4G4Vzs4w5x4HwU6MIU1B4G4UKMpjtAYG4ULsX89A4HwUSDXc9w4G0VS06jtA4E0Vrzs880sE7UU"
; , {x1:10, y1:200, x2:1060, y2:900}, data.yyyy, data.mm, data.dd, data.hh, data.min) ;起始时间:
Send {Esc}
FillDateTimeField("|<>*163$88.4TW20zwNjU0Eby0E28820HzWT11087s8V4816O9A4E0U40jYMU4TjYzxDm4FyGzW0FaWH14V8Ty1D1Dz7u9x4HwU6M4U0U4NcYoF8G1NXKzW0FayH94V85y8m281Dy9A4HwUSDXc8U498bkF825s0OzXzlb2H1408Rzy228104s0sE7UU", {x1:10, y1:200, x2:1060, y2:900}, data.yyyy, data.mm, data.dd, data.hh, data.min) ;起始日期时间:
Send {Esc}
; FillDateTimeField("|<>*164$60.AE040049zU8TU41w440UElV41A4E0ULP141DzHwVxA17tA4G4V8T141x4HwUFUl41B4G4UwQ141AYG4U03141A4HwUCk141w4G0VkS141A4E0V01Dzw0sE7UU"
; , {x1:10, y1:200, x2:1060, y2:900},	 data.yyyy, data.mm, data.dd+1, "05", data.min) ;终止时间：
Send {Esc}
FillDateTimeField("|<>*164$88.AE040zwNjU0Eby0Vy0E20HzWT11084AMF0816O9A4E0ULP140U4TjYzxDm7ok4TW0FaWH14V8G7kF0Dz7u9x4HwUFUl40U4NcYoF8G3lk4E20FayH94V800kF081Dy9A4HwUCk140U498bkF8271s4E3zlb2H1408E0Hzz8144s0sE7UU", {x1:10, y1:200, x2:1060, y2:900},	 data.yyyy, data.mm, data.dd+1, "05", data.min) ;终止日期时间：
Send {Esc}


; ********************************
; MsgBox,4096,呼吸机时间

MouseClick, left,  1150,  574


; 术后机械通气时间≥24小时：
Text:="|<>*168$171.0q01sns854TsE0012Ts03k60M00M06MTs6F10YF63zns88100F0k30y300k203u8zzm70U0H1408U0AC0M6EM7zyE06F10k3zDzWTybt301XkHMnzl0k3zwm8BKiG/00H14V8608K6N6EM8D0E0DF3OIHzDz3u8bt0Q34kXAyH02o201u8NHWG808HF4V860la8MaGM0aELzSF5TsHz012N8Y930MDx36m/0MlaUMm99G2G809H14z8U206EMKEM666Y36F9+mGT01fs8Y100k0k32y310k4zso99Rr0007H1408zbw60M6EM861437UlGA7zU0k1sUD00000S00S0U"


if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
	; FindText().Click(514, Y, "L") ;点击 是
	FindText().Click(600, Y, "L") ;点击 否
}


;ASD-4：围术期抗菌药物应用情况
Text:="|<>*157$158.Dz7zsEUDzbzxsU842208E2414020E307zm488EG4290UEzzzzxF00zw1o220V2Q45TwmE9zUV08EIzU813YkbyDzX7tF0PzjU1zy4M7ecHzl12N8W4/y2QE6F0U0E0W7uIe400006G8V214Z7uYECy4y9+2Y+UTznzkbyDzUV9N28468V02SUVYc040U488W480GGEWTuW8TzV69lG0Fy812O0V204Yg8UE8W4S8UWoYUAE20EVUEEU3Vm2842MZs+S4VG8LY0zw8y448VaF1210Y9E0U08F277z812ks1SNUoXkjzu3bzsw64b0000000000000000000000ED1008"

if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
	FindText().Click(508, Y, "L") ;点击是
}

Sleep,100
; ********************************

Sleep,300
; 预防性抗菌药物选择：
Text:="|<>*164$154.88000F80+220004G163v024040yz0014k0UDjrzsFA3M1DvzzTzZ5U00AFDzxFM0034EZ0B0UV010WG000Vz088YU008TnY0jmTzXzkzw006y1yWDz001jU2E299U20E0AE00s86OE3400C2010AYay9zz0lDzyUUNd0AE0088DzxnqMkVVUzw00221ysDz000UU0EC8FzyDu7zs008AKNZzy002343U8j6S8QQ1kU00UH0gEQ80084kNUW1SqXxcPS0021dxf6rjzwUOC3W8BX2AHAA00081U8P300020NU2szbzt72U" ;第一代或第二代头孢菌素

if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
	FindText().Click(X-97, Y-2, "L") ;点击第一代头孢
}


Send, {PGDN}
Sleep,300
MsgBox,4096,点击继续 填写首剂抗生素时间
Sleep,300
; 使用首剂抗菌药物起始时间：
FillDateTimeField("|<>*158$171.8E7zkEU01440EU48280XwEE00VDw1zwV2240U8UEzzzzxF040W21s440UEE48Hzyzt4zkEU48+Tns4E290W042Tszy0U09/s0TzV61ue40jYNDzHwVmF48FzsO940204EzGZEVwZz90WEYCG8V2810l8vsHsYc+Eezk4w5t4HwUHz7zlztldAF204x139EE0Y098WEY228V28129+W8TzV69lGG05jt8YG4UHE48FzsF94F2D4EFOGGR4N190WTY2611281288a9S2bV8IWGDXc9s4G0VFw88Fzsl14V+040128Hk0pz90W04+ks1S8149vcCTzXkMGQrzs880sE7U00000000000000US2000000000004", {x1:10, y1:200, x2:1060, y2:900},data1.yyyy, data1.mm, data1.dd, data1.hh, data1.min)
Send {Esc}


; 手术时间是否≥3小时：
Text:="|<>*156$135.07010049zXzlzy00Q0E00E3z002D0UU4E20M004k20y200000184E0Xzk7E8020E4EE000zz9zuTYE279Uk0EWEXzXTz0A184G4Xzl121U64H4EEM003kD8WTY0000063VW8yW0000V194G4bzwzw1U68FYGE7zy8494WEY0U40Uk0O24W200060N84HwV7sU4803EEoEE000U1D0WE4MU40U00E22y230U00184E0bY0zwDsw0E4EEMw0000720xXzY0U000S00S0U"


if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
	; FindText().Click(514, Y, "L") ;点击 是
	FindText().Click(600, Y, "L") ;点击 否
}

; 术中出血量是否≥1500ml：
Text:="|<>*167$162.0000000003zk0000000000000000q0A0300U30EzwTzU013sS7U01U0n0A0X41U3zkk40k00720nAk01U0k7zsX4Dz000zw3w20D20V8E01UzzoA8X4997zwk4Cn1U320V8HRlV0k4A8zw993AEzwEkkM321V8PbNV1s4A834993zk00000C33lV8P69U2o4A936993AHzzDz0M30NV8P69U4m7zt36993zk10811U30AV8H69UMlYA936990A0Fy812030AV8H69UkkkA136997zsl0810030MnAn69V0k0A1zy990A1t0Dz3y33kS7X69V0k0A006zzzzz7z8100000000000U"


if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
	; FindText().Click(514, Y, "L") ;点击 是
	FindText().Click(600, Y, "L") ;点击 否
}

MsgBox,4096,点击继续 填写术后是否抗生素及术后抗生素停止时间
Sleep,300
; 术后是否使用抗菌药物：
;
Text:="|<>*155$143.0803lzszz241zwEE120EU8U004Dw20E307zm48UEzzzzxF0000E07zUCUEE48FDw48122bw7zsU0813YkbyDzjU1zy4M7ecEM1zyTy88H94EV40204EzGZEVs20000006G8V2Cy4y9+2Y+U48403zyTy4zlzwl480Ho4AZ0E8/z040U488W4+W8TzV69lG30AI22Dl08HE48F4EXl44KYY40984AE20EVUEEWMZs+S4VG8E02TswU7zV7kUV4V+040128EU08UH7z812ks1Su3bzsw64b000000000000000000021s801"
if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
	FindText().Click(514, Y, "L") ;点击 是
	; FindText().Click(600, Y, "L") ;点击 否
}
Sleep,300

; 术后抗菌药物停止使用时间：
FillDateTimeField("|<>*158$171.0A03l1048120W08E080V0Tz00VDw00Ezk84DzzzzIE1zs107zm49s440U00401Dw48122bwE028110EV90W047zsU0y07zsFUSeWTkF09zXztDzHwV1U7zt00U14DodIm22Dn94EV90WEY8S0U0Cy4y9+2Y+aTkF0N8W49t4HwU484034EU1DEEmIE0281DwTz98WEY10UjwcW7zsFWQIWzwF088W498YG4Uk350V4EXl44KYYI0W81B0EV90WTY409849WLUdsG58WTkF08M449s4G0V009zV8GU100EW4E82817kUV90W0480284u3bzsw64b2C3zz/3U5s0sE7U00000000087UU0000000000000004", {x1:10, y1:200, x2:1060, y2:900},data1.yyyy, data1.mm, data1.dd+5, "09", data1.min)
Send {Esc}
MouseClick, left,  16,  494
Sleep, 100
Sleep,500

; 使用抗菌药物时间使用时间分层：
; MsgBox
WinWait, 郑大一单病种质控上报系统 - Google Chrome, Chrome Legacy Window
IfWinNotActive, 郑大一单病种质控上报系统 - Google Chrome, Chrome Legacy Window, WinActivate, 郑大一单病种质控上报系统 - Google Chrome, Chrome Legacy Window
WinWaitActive, 郑大一单病种质控上报系统 - Google Chrome, Chrome Legacy Window
MouseClick, left,  37,  387

Send {WheelDown 2}
Sleep,300
; ********************************
; 临床医师认为有继续使用抗菌药物治疗
Text:="|<>*158$208.8U083zw1zY4140204110U8E7zl1048120W2W00U8U20UME2E080GZ4TkzyEV427zzzze8+TtzwW0820V010Dzu/MY844124Hz120EUdzfE402TwVz04040402ZUjyHz7zrk0zz2A3pKckEE9424ZkEDzkztm4QEH94EV40204EzGZO01zwUU8GF1U1160VDwE9AYF24Rw9wGI58Jdz4A2zwV946044jy8X20UHz7zn4EU1DEEmKYYFE882YYEc0YEU8uKC2114EVIF3zw8lC+OGF9EXQ+GFmE693zU8I3zYo124F48wF15d9d994mkA996EUk4820t1lkEk88FAGw5D2Ed6bwcFc010wG360kU8S088V7kUV4V+040128GEE10zzY20E4US27WDy215Vk2xo7DzlsA9C00000000000000000000000000000087UU2"

if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
	FindText().Click(X-124, Y+0, "L")
}
Send {Esc}
Send, {PGDN}
Sleep,300

Sleep,300
;术后是否有活动性出血或血肿：
Text:="|<>*155$186.Dz7zs400401t0C04210200U02U20S40810A040013z0rkTY290W41002040G40Dz0R1zz002000U04390W4Dz7zwzwGzU813Yk80TzW030U0TazsW499020YYSYVDz448Tw0k3zwjz04aF0zw993m8YYGYV00000k41s2000UTYeF024990GEYYGYUTznzlTw24200UU44W1122990GEYYSYU0E20EE4422zkby8AWTt22993nUYYGzU4TW0ETwM1WUEY298W1122990F4YYGYUAE20EE4E0YUF42H8W1122990/4YYG41SE3zkE4004zl7yMkW11zy997qYYYG41lzm0EEw008UF420HWzw02zzk8Pzzi40000000000000000U000000000000000U"



if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
	; FindText().Click(514, Y, "L") ;点击 是
	FindText().Click(600, Y, "L") ;点击 否
}
Sleep,100

;是否有手术后并发症：：
Text:="|<>*151$129.TyDzk800C0200w844201020E30107y004Dw0V044DzkTy0u3zy0000100zz8U50020EtA20001zy800V1zzM01Ty88Ezszy0M1zy48003Tw80000A10006U800V0Dw800zzbzWzs0012101zzVUX00040U441DzwE8/z0V0G897U8z40Uzs00A0lE8482+180340U44100102G1110Vk901wU7zUU810002TsE88v280ATwU44D1s000W1410M2LzkU"

if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
	; FindText().Click(514, Y, "L") ;点击 是
	FindText().Click(600, Y, "L") ;点击 否
}
Sleep,100



WinWait, 郑大一单病种质控上报系统 - Google Chrome, Chrome Legacy Window
IfWinNotActive, 郑大一单病种质控上报系统 - Google Chrome, Chrome Legacy Window, WinActivate, 郑大一单病种质控上报系统 - Google Chrome, Chrome Legacy Window
WinWaitActive, 郑大一单病种质控上报系统 - Google Chrome, Chrome Legacy Window
MouseClick, left,  37,  387
Sleep, 100

Send, {PGDN}
Sleep,300
;常见高危因素的选择：：

Text:="|<>*157$129.00000400000000000000010EzwzzkU1zy0E1212U9zU4440UU0Dw80HzwEE4Y1487zsU47y211020E7rsbyyH0U140UU0zz80FzsWV0411kBzcW400201Tu0E4M/UU8z1844EXzwHw8EHzyWF5ztw41zUX4E0WEV32667n8WEs40004MWTYG48oFyEW14G1Dw3zs51G4WHdAG374E8aG840E118+TYW1+1FyoW15VlDy+1kl2G0YE902AH7k9E08410087m0w3yDzmC4WS1zz0U0000000001020000000000U"



if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{

	FindText().Click(X+372, Y+85, "L") ;点击
}

Send, {PGDN}
Sleep,200
Send, {PGDN}
Sleep,200

;是否进行手术后疼痛强度评估：
Text:="|<>*154$186.Dz7zt280zU1k0E07U4010wzU417y480810A0W8407y004Dw3zwzz0UXzwU0480Dz0R0jy800000080200U00UWEUY2/zk813Yk28G0001zy802DljyQzXzs44881Dz44/W85zrzk30DzqMlV8042EXW0M8100000W8A20007U80370jyVzWTUU0s80TznzkjzM20008E806BVcWx4W00jz9zU0E20EW802DzwE8/z/USjy54XzkU090U4TW0EW802001U6+12C0cW5zYEUU090UAE20EY80200102G121Ujy44Y/0c090VSE3zlE00208000Hz4Q18W4DYD0k09zVlzm0EDy0S1s000W141U8iRsNUMU090UU"


if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{

	; FindText().Click(514, Y, "L") ;点击 是
	FindText().Click(600, Y, "L") ;点击 否
}

;是否进行手术后镇痛治疗：
Text:="|<>*156$158.Dz7zt280zU1k0E07W1010EE04020E308W101zU013z0byDzn400U0zw1o2zsU000000U0S4+00247zk813Yk28G0001zy800DtjyF0c00Hzl12sW1Txzw0k3zzm2MG2ztjy4000028Uk8000S0U08zWzs01030TznzkjzM20008E80289cW1za1U040U48W00Xzz422znvyfzWE+UU0Fy812MU080060Mc48UWW8Y2880AE20EY80200102G13zwjyF0W20LY0zwI000U20004zklAG8YTtUU77z810zw1s7U0028490kWt42Fs0U"



if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{

	; FindText().Click(514, Y, "L") ;点击 是
	FindText().Click(600, Y, "L") ;点击 否
}


