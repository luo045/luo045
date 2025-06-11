SetKeyDelay, 100
#SingleInstance force
#Persistent


 MsgBox,4096,当界面出现手术开始时间和结束时间时继续
 getdata(){
     
; Sleep, 100
; MouseClick, left,  291,  299
; Sleep, 100

; Send, {PGDN}
; Sleep,200
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
     
; Sleep, 100
; MouseClick, left,  291,  299
; Sleep, 100

; Send, {PGDN}
; Sleep,200
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
    
    Send, {TAB}
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
    ; Send,{Left}	
    if (Targetmonth != 0) {
        direction := (Targetmonth > 0) ?"Right" : "Left"
        loops := Abs(Targetmonth)
        SendInput, {%direction% %loops%}
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
	SendInput, {%direction% %loops%}
}
else
	sleep,300
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

sleep,300
send,{Esc}
data1:= getdata1()
; data2:= getdata2()
send,{Esc}
FillDateTimeField("|<>*174$51.300000000808EA7W30U111XaEM60Pzgk230k311Y0EMD0s8AU2318D11Y0EMNUPzAU2324311Y0EMkkM8Ak32A3311X6Qn0AM8ADVwk0Xzw0004", {x1:10, y1:200, x2:1060, y2:900}, data.yyyy, data.mm, data.dd, data.hh, data.min) ;入住ICU：
Send {Esc} 
FillDateTimeField("|<>*163$51.0U7ztUwETzy48AQm3/50V1W0EN7848AU23/50V1Y0ENztzzgU230U0V1Y0EPzw48AU23F8UV1a0EGTY88AMnaE0W11Vw7W0MU80000U", {x1:10, y1:200, x2:1060, y2:900}, data.yyyy, data.mm, data.dd+1, 10, data.min) ;离开ICU

MsgBox,4096,输入体外循环时间 格式为
inputbox,cpbhh1,体外循环开始 hh
inputbox,CPBmin1,体外循环开始 mm
inputbox,cpbhh2,体外循环结束 hh
inputbox,CPBmin21,体外循环结束 mm
MsgBox,4096,手动返回体外循环输入框界面 再点击继续
send,{Esc}
	FillDateTimeField("|<>*166$144.48110ECyzV7sUUDz6Ps049zU48110rU84108UU81Dy9w440UDznt1YE8A7s8V4816O9A4E0U8Q291Lz8810/t6817vtDzHwVMQ49UoEyM17t/y816O9A4G4VseB9Nby8SDy1D1Dz7u9x4HwU8e9lBb28x1a580816O9B4G4U990l0jy9hZaBPy816PtAYG4U990V0j2DAZy8m281Dy9A4HwU+yVV0f2MALXsu2812G9w4G0VA8H10fyUA7U1jyDz6Q9A4E0V88410n20ABzy2281A4s0sE7U880000000000000000000000U", {x1:10, y1:200, x2:1060, y2:900},	data.yyyy, data.mm, data.dd, CPBhh1, CPBmin1) ;体外循环起始日期
FillDateTimeField("|<>*163$172.48110ECyzV20zyDz6Ptzy8800Eby0EU442S0UE7zm48U4zsUV0UUT11082zwyEN4230UU8EW0FaW2424FA4E0U8Q291Lz882zszy817vs8EyFYzxDm5VkEa2F3tUP8W48U4NcUV1/yH14V8S+VGK9zW7VgW8EXzlyWzzow5x4HwU8e1lBb28x2zszy816O88EG04oF8G0YY342zsao88W48U4NjUV1PyH94V82GE8E+kXn8aU8EW0HzW24389A4HwU+yV10f2MA2C112812G8EECUbkF824kVA42ju0k8y448zwNkW11fyH1408G210EAkU30gD0LW0E1CE48880sE7U88000000000000000000000000002", {x1:10, y1:200, x2:1060, y2:900},	data.yyyy, data.mm, data.dd, CPBhh1, CPBmin1) ;体外循环使用日期开始时间：
	FillDateTimeField("|<>*163$88.4TW20zwNjU0Eby0E28820HzWT11087s8V4816O9A4E0U40jYMU4TjYzxDm4FyGzW0FaWH14V8Ty1D1Dz7u9x4HwU6M4U0U4NcYoF8G1NXKzW0FayH94V85y8m281Dy9A4HwUSDXc8U498bkF825s0OzXzlb2H1408Rzy228104s0sE7UU", {x1:10, y1:200, x2:1060, y2:900},	data.yyyy, data.mm, data.dd, CPBhh1, CPBmin1) ;起始日期时间：
	FillDateTimeField("|<>*166$144.48110ECyzW1030Dz6Ps049zU48110rU842Tzzz81Dy9w440UDznt1YE8A4V030816O9A4E0U8Q291Lz88D1030817vtDzHwVMQ49UoEyM2Tszy816O9A4G4VseB9Nby8S600X6Dz7u9x4HwU8e9lBb28x7jsX6816O9B4G4U990l0jy9hg88XQ816PtAYG4U990V0j2DAU887U81Dy9A4HwU+yVV0f2MAHc8/E812G9w4G0VA8H10fyUAADsnADz6Q9A4E0V88410n20A089X781A4s0sE7U880000000000300000000000U", {x1:10, y1:200, x2:1060, y2:900},	data.yyyy, data.mm, data.dd, CPBhh2, CPBmin2) ;体外循环结束日期时间
		FillDateTimeField("|<>*163$88.AE040zwNjU0Eby0Vy0E20HzWT11084AEF0816O9A4E0ULP140U4TjYzxDm7ok4TW0FaWH14V8G7kF0Dz7u9x4HwUFUl40U4NcYoF8G3lk4E20FayH94V800kF081Dy9A4HwUCk140U498bkF8270s4E3zlb2H1408E0Hzz8104s0sE7UU", {x1:10, y1:200, x2:1060, y2:900},	data.yyyy, data.mm, data.dd, CPBhh2, CPBmin2) ;终止日期时间：
FillDateTimeField("|<>*158$171.8E220UQxz241zwTy47l0U0U00VDw120EE9s210Tz8EW0FzW9zzzxs440ULz7m38UEM44124E244GEU0U90W04270WE5zm20byDzW0Ewyw43zlDzHwVks8H18UwkAYF24E244FDwEW90WEYC+VGK9TW7VYW8EXzkwWM024Ft4HwUFI3WP84Fe4zlzwE244GrwTy98WEY2GEAE9TWNEUW8EW0EUykUUS18YG4UGG121E4S94o124E2TwE444c90WTY2jcEE+0a10VUEEW0EYWCUX4ls4G0VMEY21Hx084T224TyAcK7wUVd0W04+210E80U10gC0LW0E1C0UU400sE7UEE000000000000000000000000004", {x1:10, y1:200, x2:1060, y2:900},	data.yyyy, data.mm, data.dd, CPBhh2, CPBmin2) ;体外循环使用日期结束时间：

send,{Esc}
MsgBox,4096,呼吸机时间


  MouseClick, left,  1150,  574
Sleep, 100

	FillDateTimeField("|<>*164$60.4TW20049zU40W21w440UTUW4FA4E0U40jYNDzHwV4TYjtA4G4Vzs4w5x4HwU6MIU1B4G4UKMpjtAYG4ULsX89A4HwUSDXc9w4G0VS06jtA4E0Vrzs880sE7UU"
, {x1:10, y1:200, x2:1060, y2:900}, data.yyyy, data.mm, data.dd, data.hh, data.min) ;起始时间:
Send {Esc} 
	FillDateTimeField("|<>*163$88.4TW20zwNjU0Eby0E28820HzWT11087s8V4816O9A4E0U40jYMU4TjYzxDm4FyGzW0FaWH14V8Ty1D1Dz7u9x4HwU6M4U0U4NcYoF8G1NXKzW0FayH94V85y8m281Dy9A4HwUSDXc8U498bkF825s0OzXzlb2H1408Rzy228104s0sE7UU", {x1:10, y1:200, x2:1060, y2:900}, data.yyyy, data.mm, data.dd, data.hh, data.min) ;起始日期时间:
	Send {Esc} 
	FillDateTimeField("|<>*164$60.AE040049zU8TU41w440UElV41A4E0ULP141DzHwVxA17tA4G4V8T141x4HwUFUl41B4G4UwQ141AYG4U03141A4HwUCk141w4G0VkS141A4E0V01Dzw0sE7UU"
, {x1:10, y1:200, x2:1060, y2:900},	 data.yyyy, data.mm, data.dd+1, "05", data.min) ;终止时间： 
Send {Esc} 
FillDateTimeField("|<>*164$88.AE040zwNjU0Eby0Vy0E20HzWT11084AMF0816O9A4E0ULP140U4TjYzxDm7ok4TW0FaWH14V8G7kF0Dz7u9x4HwUFUl40U4NcYoF8G3lk4E20FayH94V800kF081Dy9A4HwUCk140U498bkF8271s4E3zlb2H1408E0Hzz8144s0sE7UU", {x1:10, y1:200, x2:1060, y2:900},	 data.yyyy, data.mm, data.dd+1, "05", data.min) ;终止日期时间：
Send {Esc} 
MsgBox,4096,抗生素时间


  MouseClick, left,  1150,  574
Sleep, 100

FillDateTimeField("|<>*158$171.8E7zkEU01440EU48280XwEE00VDw1zwV2240U8UEzzzzxF040W21s440UEE48Hzyzt4zkEU48+Tns4E290W042Tszy0U09/s0TzV61ue40jYNDzHwVmF48FzsO940204EzGZEVwZz90WEYCG8V2810l8vsHsYc+Eezk4w5t4HwUHz7zlztldAF204x139EE0Y098WEY228V28129+W8TzV69lGG05jt8YG4UHE48FzsF94F2D4EFOGGR4N190WTY2611281288a9S2bV8IWGDXc9s4G0VFw88Fzsl14V+040128Hk0pz90W04+ks1S8149vcCTzXkMGQrzs880sE7U00000000000000US2000000000004", {x1:10, y1:200, x2:1060, y2:900},data1.yyyy, data1.mm, data1.dd, data1.hh, data1.min)
Send {Esc} 
; MsgBox,4096,术后医嘱-是否使用抗抗生素 抗生素结束时间
	FillDateTimeField("|<>*158$171.0A03l1048120W08E080V0Tz00VDw00Ezk84DzzzzIE1zs107zm49s440U00401Dw48122bwE028110EV90W047zsU0y07zsFUSeWTkF09zXztDzHwV1U7zt00U14DodIm22Dn94EV90WEY8S0U0Cy4y9+2Y+aTkF0N8W49t4HwU484034EU1DEEmIE0281DwTz98WEY10UjwcW7zsFWQIWzwF088W498YG4Uk350V4EXl44KYYI0W81B0EV90WTY409849WLUdsG58WTkF08M449s4G0V009zV8GU100EW4E82817kUV90W0480284u3bzsw64b2C3zz/3U5s0sE7U00000000087UU0000000000000004", {x1:10, y1:200, x2:1060, y2:900},data1.yyyy, data1.mm, data1.dd+5, "09", data1.min)
Send {Esc} 

MsgBox,4096,术后医嘱-是否使用抗血小板药物：首剂用药日期时间：
FillDateTimeField("|<>*155$110.4800Fzw487zV1w024zl20E4EVTzt08zlD0UU7zyzt48EEUE244GE8U1100GFzw8k40VtwbzdyLzVcYEV4Dl08EF90WEZ086948GI4Ty7YHm8btTyAOFzwx140V14YW92I0V4YEV2AF08ET94WEZzsF948F14E2TwGE8btE24EG24w940V94w290LzX44UV00FzsUV90W0508VD0/lsAE209k1kUD0000000US0000000008"
, {x1:10, y1:200, x2:1060, y2:900}, data.yyyy, data.mm, data.dd, 23, data.min) ;手术当日23点
Send {Esc} 
; F4::
; getdata2()
; data2:= getdata2()
; data := getdata()
; return


ExitApp

