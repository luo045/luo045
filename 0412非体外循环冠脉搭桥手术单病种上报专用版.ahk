﻿﻿﻿#SingleInstance Force
SetKeyDelay, 200
#SingleInstance force
#Persistent
SetWorkingDir %A_ScriptDir%
SetTitleMatchMode, 2
CoordMode, Mouse, Screen



 ; ********************************
; 查找 主要诊断ICD-10四位亚目编码与名称：
Text:="|<>*170$128.EA4kBUwT00A73zw8kTznzm71hwMtYQ0738aH460F0k4H8Lc6M11U3kX9YlTw4EA11ZYu1Y0E80AMGNAk094nzyWByUN0420364aHAUWF8k4X8HDqE10U0lV94r88oGA19YZmFY0E9wAMGFAn455Xzm6NyYN0420364gTAF1FEk4W8PN6M11U0lXC0n4kQQA1+A4mFX6Fk0AAW0Al81430HC1z4MT7k031kU3A20F0zw400F00000000DznzzzzA100000000000020A0000008"
;诊断ICD-10四位亚目

if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
	FindText().Click(X+498, Y, "L")
}
Sleep,50
; 鼠标依次诊断ICD-10四位亚目输入框
SendInput, {DOWN}{DOWN}{ENTER} ;选择125.1动脉硬化性心脏病
; ********************************


 ; ********************************
; 查找 诊断ICD-10六位临床扩展编码 与名称：
Text:="|<>*169$187.862M6kSDU063U404M140208A7zkUUzz271hwMtYQ073810466a01U42208bz40UaEjEAk2307V6002zvHyTzjztzwG0W8E6KHc6E10U0lV00301eE80180U2Lzl4FoFjo380UE0MkjztY4q4440Y0F4CU1288m4nxY0E80AME01m2O03zwO0DzWTwxz4mGt8m084y6A88UNWBTt3Ut04F3Aem0m6NyYN0420364A8AF6hYWc4U3zxyJN0F4EqmAk2301X6A669XKmGG2E1/43zYzsckH96AN700km41X4VfN+8V80YwDZGEANkDsX3sy00MC40NUEpzi48c0Xg9Gds61004E0000000404zz2kI21o0F0kcQYSU"
;诊断ICD-10六位临床扩展编码
if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
	FindText().Click(X+498, Y, "L")
}
Sleep,50
SendInput, {DOWN}{ENTER} ;默认125.103冠状动脉粥样硬化性心脏病
; ********************************


 ; ********************************
; 查找 主要手术操作栏中提取ICD-9-CM-3四位亚目编码与名称：
Text:="|<>*166$251.9z140l40k2Dvy0MD7k0700yM303Vzy487zszw88DzkH22E1W81U4En/wln8M0H034s609W94881Y108Lz20Xzw8zzUVzyyzaK9a0EM1W0A1kQ014G8jy3820FA28F100GE6Tu64F3DYH80UE260E3Uc028YF00aF7zYzwEW2RtgUA04A8vwN9aE10U4A1U6XE0AF8aEFga81Dk0XADeLFwQ08MHU0yGAU210AM30B4U3kWFQUVN8E25znrwtrW21s0EkizxYoN042TDrq0P9DUl8W8V2mEzwHeaU8EE447bwzz4238km084010A0mW00mlwF47Z109zJR0Ejz8Dn01328bbxVa0EM020A1b401b08X87+20EDyPzd3UEE60060H8QbX6AVU0A0QH680a80F2k6E40XpIo12+EUUA00A0jk1Na7lw03U0D68E1sE0W10AUDzBeds2R4N10PzUM7Fy3V00000000000000zz5zzzyE22FnEs000000000000000000000000000102000000000001"

;主要手术操作栏中提取ICD-9-CM-3四位亚目编码与名称：
if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
	FindText().Click(X+498, Y, "L")
}
Sleep,50
SendInput, {DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{ENTER} ;36.15单乳房内动脉-冠状动脉旁路移植
; ********************************

; ********************************
; 查找主要手术操作栏中提取ICD-9-CM-3六位临床扩展编码与名称：
Text:="|<>*167$279.9z140l40k2Dvy0MD7k0700yM303U204814030847zsUUzz1A89068U60F3Ajn7AVU1A0AHUM0a0E0UUcU0810Uk35zkU8zz2Dzs8TzjjtZWNU460MU30Q700E10/zZDxzyyzrztA28F100GE6Tu64F3DYH80UE260E3Uc0200100fEA0140k3Hzl289ram0k0EkXjlYaN0420Ek60OB00nzzN15F1VU+U6F3w08n3uZoT70264s0DYX80UE360k3F80w00788c0DzlY0zz5znrwtrW21s0EkizxYoN042TDrq0P9DUk8E8V5DtVUMU6F1CeO0V10EESTnzwE8AX380UE040k3+8033214Ed9AC5A0zzTpLE4/zm3wk0EkW9tzMNU4600U30Nl00ME88m599Zc904YEDyPzd3UEE60060H8QbX6AVU0A0QH680a40V2kd9/Al80YwDJHE48d220k00k2z05aMT7k0C00wMV07V06845DtFW/04xXOeS0bF6EE6zs61oTUsE000000000000080FTy91EA7E1632FnEwU"

;主要手术操作栏中提取ICD-9-CM-3六位临床扩展编码与名称：
if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
	FindText().Click(X+498, Y, "L")
}

Sleep,50

SendInput, {DOWN}{ENTER} ;36.1500单乳房内动脉-冠状动脉旁路移植




; ********************************
; 是否为出院后31天内非预期重复住院：单项选择是否/男女：
Text:="|<>*154$114.Dz7zsW00U7W00SD00zy810A0G08V400zk11U10Dz0R0208V5TsU01XU10813Yk208V5E8U01UU10Dz449zyDz6Dszz10Vzz00000220V500U0C0U10Tznzk22EUZzsU01UU200E20E4WEUY4Ujw0UU2U4TW0EAGEUZ4Uc40UU4EAE20EM2EUb4Z84FUU88SE3zkk6TzY8ZDwD0Uk6lzm0F0w00Yku8400103U"
;是否为出院后31天(问题的前8个字)：
if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
	;FindText().Click(514, Y, "L") ;点击 是
  FindText().Click(600, Y, "L") ;点击 否
}
Sleep,50
; ********************************
SendInput, {PGDN}
Sleep,300
; MsgBox,4096,,确认手术时间正确 再继续
MsgBox,4096,确认手术时间正确 再继续
 getdata(){
     
Sleep, 100
; MouseClick, left,  291,  299
; Sleep, 100

; SendInput, {PGDN}
Sleep,50
; 查找 手术结束 复制并发送剪贴板
Text:="|<>*163$56.070BUUE0k7y0349zzzw300k4V0300kDzzkEDz7zs309zWAE301s600X40k0h1vyDzDzwH8kUUS030MlU88/E0kAAAu2An0A030kzYAAS00k08830U"


if (ok:=FindText(X, Y, 153-150000, 641-150000, 153+150000, 641+150000, 0, 0, Text))
{
	FindText().Click(X+520, Y+4, "L")
}
Sleep, 100
SendInput, {Esc}{CtrlDown}a{CtrlUp}{CtrlDown}c{CtrlUp}
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

; SendInput, {PGDN}
Sleep,50
; 查找 手术开始  日期时间： 复制并发送剪贴板
Text:="|<>*165$56.070BVzy887y03A24220300k0V0V40kDzw8EyFbzs30244js301sDzxD10k0h08EG0DzwH8245js30MlUV0m20kAAAEECUUA030846zsS00k41222U"


if (ok:=FindText(X, Y, 153-150000, 641-150000, 153+150000, 641+150000, 0, 0, Text))
{
	FindText().Click(X+520, Y+4, "L")
}
Sleep, 300
SendInput, {Esc}{CtrlDown}a{CtrlUp}{CtrlDown}c{CtrlUp}
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
Sleep,30
; 鼠标依次点击年,月,日,时间
; 点击月份输入框箭头,激活月份输入框
Text:="|<>*144$28.Dw000UE0021000Dw07wUE0DW100QDw00UUE0041000E40021k008"
if (ok:=FindText(X, Y, 555-150000, 629-150000, 555+150000, 629+150000, 0, 0, Text))
{
	FindText().Click(X, Y, "L")
}
Sleep,300
    ; Send, %mm% ;错误
  
SendInput, {TAB}
Send, %yyyy%
SendInput, {TAB}
Send, %hh%
SendInput, {TAB}
Send, %min%
SendInput, {Esc}
SendInput, {TAB}
   ; ********************************这部分代替不能直接数字输入月份的办法
Currentmonth :=A_Mon+0
Targetmonth :=mm
EnvSub,Targetmonth,%Currentmonth%
    Send,{Left}
    if (Targetmonth != 0) {
        direction := (Targetmonth > 0) ?"Right" : "Left"
        loops := Abs(Targetmonth)
        SendInput, {%direction% %loops%}
    } else {
        Sleep, 300
    }
    ; SendInput, {Enter}
    SendInput, {Esc}
    
    
    
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
SendInput, {TAB}{Esc}
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
Sleep, 100
; MouseClick, left,  18,  888
Sleep, 100
SendInput, {PGDN}
Sleep,300
; CABG-1 术前评估部分
; 1********************************
; 是否实施手术前的冠状动脉造影评估：单项选择
Text:="|<>*165$140.0U2400Q0q0V0EUTzV3E0ESM040H0Ts0AkAE88408EWTY4UVzyzzkA030zzrrs04I8011/kE0WM030zzk0152TV7zw1zS4VW0UETzUA1yGFUU7wkk04oVM4UCpUA07UEYYG8044A7t9DQA86j8302o7t9wmTt330EGGK0a1jGzzlAV2GF4WGFlc8AbZnzzGrUA1X6TYYE8Y5oG2G9+I0k4h030kko99429ClAFQWHYUvl/10k0A122T0WE8K6MsYlBk6iTVs030Hb4Hl7y70UAulk000000000000000000020008"

;实施手术前的冠状动脉造影评估：：
if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
	FindText().Click(514, Y, "L") ;点击 是
  ; FindText().Click(600, Y, "L") ;点击 否
}
Sleep,50
; ********************************

; 2********************************
; 冠状动脉病变数量：下拉选项
Text:="|<>*166$115.00000000000000007zUTzV3E0ESM04020GY30E80EV4z8911zy0U1W1zs00FEU044j2U0zzzvw003w8zzUDvkYzzUY0l4zzW1zAA01B8KMA0GVwWAl102263wabi4639BZu7zUTt330EGGK6zt02A53AE4YXXEEND/ZFYTwTnVzs2ELF898YdEdu244EU3019q9W/YGQYL70w1sszz2Y25VaC9AHG0Uz0wq0k1Xz3UE6RMs13nUzUVzzs00000200000000000008"
;冠状动脉病变数量
if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
	FindText().Click(X+498, Y, "L")
}
Sleep,30
SendInput, {DOWN}{DOWN}{DOWN}{ENTER} ;选择3支
; ********************************


; 3********************************
; 血管病变主要位置：单选
Text:="|<>*158$116.0U220100U081zy487zs0E1zzDzk4030181114W0zw2m+01zz0E7zsjyTzU997ztzy187zt4W80000GGF02M00G010F8aENzz4YYTyW01Ya0E7zva40009920Vjy00U40408V3zk2GEzseAXzVzyzzm8UU40YY802bc8E0E110V8Dz0993zkeC1s040DU8G20EGGEU4G0Uy0101w210U47zz3y0VtkSzzrUMjzzzkU"
;血管病变主要位置
if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
	FindText().Click(X+377, Y+43, "L") ;第1项 左主干LM病变（病变狭窄超过50%）
  FindText().Click(X+377, Y+86, "L") ;第2项 LAD
  FindText().Click(X+377, Y+127, "L") ;第3项 LCX
  FindText().Click(X+377, Y+170, "L") ;第4项 RCA
}
; ********************************

; 4********************************
; 手术前的危险因素评估：单选
Text:="|<>*152$149.0DUA0M6360k00M7zw30M00lU0zk0P0MM6A1znskA0PzwPz1X00300n7zyMTa36mkMkkA0kk3600601U001xXMABgklVbzU1UBzk7zvzzDsnC7Tzyklzv0k0PMMM00M0D0MxaMAk0yyn67zzyKlkk00k0j0lvAqNjtg06AAMkAj5VUC1U1S1zqTanMnNaAwNz0MM3TsTzz6q37gnBalalgNgkMkzz6kk060Na6DNa3BXBvMqBjzlVUBVU0A1X6DynA6PQSNVsP0lXn0P33UM666MlaMBa3kn306Ng760q670k0A0lXDkPA7UA7zxXAAA1jw07U0M1jSNXgDzDzA0SSA0M3MM2"

;手术前的危险因素评估：
if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
	FindText().Click(X+461, Y+39, "L") ;是否使用“SinoSCOREⅡ风险评估表”进行手术前风险评估： 选择否
 }
; ********************************
SendInput, {PGDN}
Sleep,300
; 5********************************
; 术前最后一次超声心动图左室射血分数(%)：
Text:="|<>*158$263.0A101zs0D000408z0804008Dzk801U20U402E9G0X140042420Ezk00480EWTzk41yEF0UE1zyT10E08k243944001zzbzV00004znt40U0800UbxDzu04XzbzUUkzDY2E87zs00812000010V28zy2907tAG203zlw49960kMW8YUMEM3wZzzbzs000H2/U00I202Gb4401U2+8GGM0Nt4V+0EVs498W080000EYz03zl86TYYC880467mEYYDyAgF1ZkU48DmFzwE0Tzt100z48WE489D3kzwTy8YV982462W0IV0E8EYW+Ejw0025112Dz4U8UmFUV2000z12GE08zb418W30Az97kVE800892S4E2F0994Ul441zs624YUEEW482F44092G8a4UE00EW47sU0204m9C288000I4990UUwQM8YME0244zi9zU0063D8G0044AQG3YUE00188GG231dAEUsUU04tk3XW1000E2rzk00Ds0HbzsTzzzxnnzzAwA4AU0100000000000000000000001080E000000000U0000U041"
;手术前的危险因素评估：
if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
	FindText().Click(X+461, Y, "L") ;术前最后一次超声心动图左室射血分数(%)
 }
 send,52
; ********************************


; 6********************************
; 左室舒张末内径(mm)：：
Text:="|<>*167$137.200M0nzxV0M060/z200004040TzXMM/40k0A0kAA000083zyU1APUKkzzDzn0kE000080U0zwzW7h030MlYXVWtlQsEl00M08zsM060l32Mm6QnCMlW011Xy9lzrztX6/0oAFa8lUDz7zUUHxc0M3DAk08EX8FX0FU0k7sU/E1s6lObyEV6EX6130Tz8l0KEBgC1l1UV2AV6A26030FW1hkX4M1W31W4N2AEsA060z42Qm66k346148m4MVbzzzx6sQUUA1UQ/zn00002000000000000000002000041"

;左室舒张末内径(mm)：：
if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
	FindText().Click(X+461, Y, "L") ;术前最后一次超声心动图左室射血分数(%)
 }
 send,49
; ********************************

; 6********************************
; 左室舒张末内径(mm)：：
Text:="|<>*167$95.2008060/z200004040DzUA0kAA000083zyE1Dzn0kE000080U0zyMlYXVWtlQsEl0104l32Mm6QnCMlW0221X6/0oAFa8lUDz7zvDAk08EX8FX0FU8U6lObyEV6EX6130nyC1l1UV2AV6A26144M1W31W4N2AEsA2M8k346148m4MVbzt3lUQ/zn0000200000000002000041"

;左房内径(mm)：
if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
	FindText().Click(X+461, Y, "L") ;左房内径(mm)：
 }
 send,43
; ********************************


;CABG-2手术适应症与急症手术指征：
Text:="|<>*175$222.0000000001w1U03k300k1U0A00M01w1U3A0M0000000003z01gAy01U0M1U0Ds0A3z01g3AsTz000000D00601a661zzDzlzwMk7zs601a3DUkMDX7sy0NU0601U661U0A01U0lU600601UDgBUMNX3Ba0NU3zxzzVztX0g01U1zwK03zxzzXADMMlZXD601U0603k061VaTzn000ADzs603k37wPMk5XD00300605sS61UqAA3zwzw660605s3k0nTk5Xv1zX00605s6zlsqAA00A0A660605s3jtnMk7XDD0607zyBg6klg6Rg00BzwCq7zyBgDAOnMkAnD60A0060Na6klaAhjjzg30Krk60Na3AMnMNgnBa0NU060lX6zlaABg00ANa6q060lX3DsnMDRzsw0TU061VVa01UMNg00BNXAq061VVXAMnM000000000601U/01UkNg00BMPAq0601U3AMnM000000000S01UNzvTzrzk0vDsPzsS01UDDszzU"
if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
	FindText().Click(X+349, Y+127, "L") ;选择 三支病变，且狭窄程度均超过70%
 }
; ********************************


; 1********************************

SendInput, {PGDN}
Sleep,300

; 是否实施急诊CABG手术：单项选择
Text:="|<>*157$155.Dz7zs208E0U10k3kM7kDU1s180E20M0208U3y13UMkk8lV7y0280zw1o7ztzz881AUU1EFW00E040108Qa80F80zy0mW04UX800UDzsnzl126821000u8o09V4E0zy0s1U00002E72lzsEY80F3sVk200U0TznzkkU+wU00YGE1W4N0U401000U40UG07d7zV38U3y8G1DzsmE0Fy817zx8Q0U28VUA4Ea20E24M1W0E20M2EUYZ561XMAX640U888rY0zw3C4Y589Ak1wUNw7k100E1szt08s1LDkTU200000000S00U0U"

if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
	; FindText().Click(514, Y, "L") ;点击 是
  FindText().Click(600, Y, "L") ;点击 否
}
Sleep,300
; ********************************
; 1********************************
; 手术切口的选择：
Text:="|<>*167$102.070BUU0TzV212k9zUTs0AkbyM1W20Yk8VU0k0A0UWM1brsbyyH00kDzwwWM1YI8Ak8C1TzUA3VWM1YM/Uk8z10k0S0VWM1YG8jzDVk0k0h0VWM1bn8WMs40zzlAUZ2M1YF8WN9zU0k6AMh2M1YE8aP8400kAAAl2M1YE8gC/zl0k0A0W6Tzbk9k08417U0A06wM1YHkDzs4000000A00000000000U"
if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
	FindText().Click(X+498, Y, "L")
}
SendInput, {DOWN}{ENTER} ;选择正中切口
Sleep,30

; ********************************

; 1********************************
; 是否合并其他手术：
Text:="|<>*151$115.TyDzk2084210U00C020810A02U2410UI0Ts00E7zUCU28DznzwGD0000020EtA420V0E89wU01zy5zsUV7zUEUDwBUHzs1U200000008E42CE8001c0zzbzVztzzXz184001200E20EU42410UYSzzl0U2Dl08E2127zyG00030A340U4811100090E01027m0Ty7zV0UAC4U8400037z8120F0EE0mDsS0000U"

if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
	FindText().Click(X+498, Y, "L")
}
SendInput, {DOWN}{DOWN}{DOWN}{ENTER} ;不合并搭桥以外的心外科手术
Sleep,30

; ********************************

; 1********************************
; 选择手术方法：
Text:="|<>*168$60.070BU30EA0Ts0Ak10MA00k0A3zzBzU0kDzw80UA1TzUA080MA10k0S0Dw8A00k0h08A3zkzzlAU8A8800k6AMMA8G00kAAAEAEV10k0A0k8F7V7U0A1XsHsU0000300000U"

if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
	FindText().Click(X+498, Y, "L")
}
SendInput, {DOWN}{ENTER} ;单纯CABG
Sleep,30

; ********************************

; 是否使用体外循环：单项选择
Text:="|<>*157$116.Dz7zsEUDzV20EE43bjs20E307zm48EU442S0UE0zw1o220V2/zXt1YE8A0813YkbyDzW70WE5zm20Hzl12N8W49VkEa2F1tU400006G8V2M+59MZy8S0TznzkbyDzW+UQHN0WBE040U488W48YY342LsaI0Fy812O0V2990V0c2D4UAE20EVUEEWjcEE+0a10LY0zw8y448kV842bu0E77z812ks1S88410U204000000000022000000008"

if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
	; FindText().Click(514, Y, "L") ;点击 是
  FindText().Click(X+621, Y+2, "L") ;点击 否 情况特殊，项目较长造成否选项右移
}
Sleep,50
; ********************************

; 是否CABG术中是否放置漂浮导管：
Text:="|<>*157$182.7UkDUT09020DyDzsUUTz5zs0CTy6206AA2AME280U20U20484WFWEHs40Xzy102UX400U7zkzs3c061zw3z2YFzs4o0U188m0Dzt248238nzy0U09248E0jzss0H28U070EV3zX220lDzv2Ek040+02C04Ey8Q0U48E00000AE20Pz4zkzwzwUU348m108124zzXzUwYDy0000M0421080zW4UEmETz080U8F+2007wE8Dzszk30M8VA4EX48EXw8243UzsE05zs0E800Mq38lV888208U20V0k805zt0U443zUnt0nsDU200U7c0zsmP2014Uk80V0U8A0000000U082Dy82/c/zyK9cS03k7w0000000000000000000008Q000000008"


if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
		; FindText().Click(514, Y, "L") ;点击 是
  FindText().Click(600, Y, "L") ;点击 否
}
Sleep,50
; ********************************

; 围术期使用血制品：
Text:="|<>*155$115.zz0211w8E7zk40I17zUEUU0FzW7zm4840+4W0E/zE00EF44124TyDuF08489zyDDWTszy994V8zw6zY1U44H94EV4YUEYE23221s3m9YW8EWGHzm000zx12114Hz7zl9849DbkEWV0UUy88W48YYToYG88FH0Bzl4o124GG9+G944/90298W6112994Y94W6440088F7kUV4YWS4yT3zy0004sgC0Ljzw8CF8UU100000000000000000E"


if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
		; FindText().Click(514, Y, "L") ;点击 是
  FindText().Click(600, Y, "L") ;点击 否
}
Sleep,50
; ********************************
MsgBox,4096,,手动填写冠脉搭桥情况后再继续 
Sleep, 100
rightsite("|<>*153$120.0k01sM063nDwM000kk000q3z0Ny63NXMM000kTz00n300Na631VkTzjUkM300k301xaSzsDwk0BzxU30zznzwNa620BhjyBUljv01s300Na6frhg00BUlgP02w300xaCfljwzyDglgP72w3zsxaDzlhg06Baljv76q3kNNaKjlhg06BalgP0An3kMNa6h1jw06BUlgP0MlakMNa6hthg06DUljv7kkqkMP7ahthw07hUlU370kAzsP7biuk003U0lU300k0kMS3a6STy01U3lUD0U") ;术后机械通气时间：移动到y=300
Sleep, 400


; 术后是否入住ICU/术后复苏室：
Text:="|<>*164$188.0000000000U000000000000000000000BU0SDz7zs404863l1UEBU0S400V03003ADw30E300U111XaEMA3ADw3zzzzTzU0k200zw3w0A0rzMU4620k201U424408DzwU0A13gk30A46E11VjzwU0zz0E0zw430Dznzl331k711Y0EME30Dzm0Fzs3011s20000000G3kEN04641s200zw1211U0h0U0zznzkAUBzaE11X0h0U0202Eszw0H8/zU40U424311Y0EMUH8/zVzlY+0k0MlaUMFy811VUkENU44MMlaUMg8H2XzsAAB86AE20EkAA46ANn4AAB86lw8UY30430HzbY0zwM1X11Vw7V030HzUzUEM0k10k8UP7z81408zz0000k0k8UNk7MwDzw00000000000000000080000000400002"
if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
	FindText().Click(508, Y, "L") ;点击是
}

Sleep,300
; ********************************

; 在ICU/术后复苏室是否实施机械通气：
Text:="|<>*163$230.1063l1UEBU0S400V030Dz7zs208E3DUUIFzV000E1XaEMA34Dw3zzzzTzX0E300E1A0m884W8kTy3zzMU4620k201U424408zw1w7zvzzyWDzw1k800206E11VjzwU0zz0E0zwA13gl029038UUM1zbzkFW1Y0EME30Dzm0Fzs303zl3368210m8BKiG9004EUN04641s200zw1211U00000G0vKSW6ocbyTy0BzaE11X0h0U0202EszwzznzkkUOwbcVZC98U0U721Y0EMUH8/zVzlY+0k040U42M6x/m8fz2Ts082kUNU44MMlaUMc8H2XzsFy81Dzx/QAWGIUYW02EA86ANn4AAB86kw8UY30AE20E30Go38YZN9DU0IH21Vw7V030HzUzUEM0k7Y0zw3C4g4o99Rr00074zz0000k0k8UNk78wDzz7z8170OtyC1WY8Dz01U000000800000004000000000000000000000008"

if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
	FindText().Click(508, Y, "L") ;点击是
}

Sleep,300
; ********************************
 ; MsgBox,4096,,手动填写冠脉搭桥情况后再继续 最后一个桥放在屏幕中间蓝白交界位置 屏幕最下方为是否预防性使用抗生素要显露出来
; Sleep, 100
; MouseClick, left,  18,  888

; MsgBox,4096,输入体外循环时间 格式为
; inputbox,cpbhh1,体外循环开始 hh
; inputbox,CPBmin1,体外循环开始 mm
; inputbox,cpbhh2,体外循环结束 hh
; inputbox,CPBmin21,体外循环结束 mm
; MsgBox,4096,手动返回体外循环输入框界面 再点击继续
; send,{Esc}
	; FillDateTimeField("|<>*166$144.48110ECyzV7sUUDz6Ps049zU48110rU84108UU81Dy9w440UDznt1YE8A7s8V4816O9A4E0U8Q291Lz8810/t6817vtDzHwVMQ49UoEyM17t/y816O9A4G4VseB9Nby8SDy1D1Dz7u9x4HwU8e9lBb28x1a580816O9B4G4U990l0jy9hZaBPy816PtAYG4U990V0j2DAZy8m281Dy9A4HwU+yVV0f2MALXsu2812G9w4G0VA8H10fyUA7U1jyDz6Q9A4E0V88410n20ABzy2281A4s0sE7U880000000000000000000000U", {x1:10, y1:200, x2:1060, y2:900},	data.yyyy, data.mm, data.dd, CPBhh1, CPBmin1)

	; FillDateTimeField("|<>*166$144.48110ECyzW1030Dz6Ps049zU48110rU842Tzzz81Dy9w440UDznt1YE8A4V030816O9A4E0U8Q291Lz88D1030817vtDzHwVMQ49UoEyM2Tszy816O9A4G4VseB9Nby8S600X6Dz7u9x4HwU8e9lBb28x7jsX6816O9B4G4U990l0jy9hg88XQ816PtAYG4U990V0j2DAU887U81Dy9A4HwU+yVV0f2MAHc8/E812G9w4G0VA8H10fyUAADsnADz6Q9A4E0V88410n20A089X781A4s0sE7U880000000000300000000000U", {x1:10, y1:200, x2:1060, y2:900},	data.yyyy, data.mm, data.dd, CPBhh2, CPBmin2)


; send,{Esc}
; MsgBox,4096,呼吸机时间

  ; MouseClick, left,  1150,  574
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

; CABG-4预防抗菌药应用时机
Text:="|<>*157$158.Dz7zsEUDzbzxsU842208E2414020E307zm488EG4290UEzzzzxF00zw1o220V2Q45TwmE9zUV08EIzU813YkbyDzX7tF0PzjU1zy4M7ecHzl12N8W4/y2QE6F0U0E0W7uIe400006G8V214Z7uYECy4y9+2Y+UTznzkbyDzUV9N28468V02SUVYc040U488W480GGEWTuW8TzV69lG0Fy812O0V204Yg8UE8W4S8UWoYUAE20EVUEEU3Vm2842MZs+S4VG8LY0zw8y448VaF1210Y9E0U08F277z812ks1SNUoXkjzu3bzsw64b0000000000000000000000ED1008"

if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
	FindText().Click(508, Y, "L") ;点击是
}

Sleep,50
; ********************************

; 预防性抗菌药物选择：
Text:="|<>*164$154.88000F80+220004G163v024040yz0014k0UDjrzsFA3M1DvzzTzZ5U00AFDzxFM0034EZ0B0UV010WG000Vz088YU008TnY0jmTzXzkzw006y1yWDz001jU2E299U20E0AE00s86OE3400C2010AYay9zz0lDzyUUNd0AE0088DzxnqMkVVUzw00221ysDz000UU0EC8FzyDu7zs008AKNZzy002343U8j6S8QQ1kU00UH0gEQ80084kNUW1SqXxcPS0021dxf6rjzwUOC3W8BX2AHAA00081U8P300020NU2szbzt72U" ;第一代头孢

if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
	FindText().Click(X-97, Y-2, "L") ;点击第一代头孢
}
send,{Esc}
; FindText().Click(X-97, Y-2, "L") ;点击第一代头孢
SendInput, {PGDN}
Sleep,300
Sleep,30

FillDateTimeField("|<>*158$171.8E7zkEU01440EU48280XwEE00VDw1zwV2240U8UEzzzzxF040W21s440UEE48Hzyzt4zkEU48+Tns4E290W042Tszy0U09/s0TzV61ue40jYNDzHwVmF48FzsO940204EzGZEVwZz90WEYCG8V2810l8vsHsYc+Eezk4w5t4HwUHz7zlztldAF204x139EE0Y098WEY228V28129+W8TzV69lGG05jt8YG4UHE48FzsF94F2D4EFOGGR4N190WTY2611281288a9S2bV8IWGDXc9s4G0VFw88Fzsl14V+040128Hk0pz90W04+ks1S8149vcCTzXkMGQrzs880sE7U00000000000000US2000000000004", {x1:10, y1:200, x2:1060, y2:900},data1.yyyy, data1.mm, data1.dd, data1.hh, data1.min)
Send {Esc} 
	FillDateTimeField("|<>*158$171.0A03l1048120W08E080V0Tz00VDw00Ezk84DzzzzIE1zs107zm49s440U00401Dw48122bwE028110EV90W047zsU0y07zsFUSeWTkF09zXztDzHwV1U7zt00U14DodIm22Dn94EV90WEY8S0U0Cy4y9+2Y+aTkF0N8W49t4HwU484034EU1DEEmIE0281DwTz98WEY10UjwcW7zsFWQIWzwF088W498YG4Uk350V4EXl44KYYI0W81B0EV90WTY409849WLUdsG58WTkF08M449s4G0V009zV8GU100EW4E82817kUV90W0480284u3bzsw64b2C3zz/3U5s0sE7U00000000087UU0000000000000004", {x1:10, y1:200, x2:1060, y2:900},data1.yyyy, data1.mm, data1.dd+5, "09", data1.min)
Send {Esc} 

rightsite("|<>*158$171.0A03l1048120W08E080V0Tz00VDw00Ezk84DzzzzIE1zs107zm49s440U00401Dw48122bwE028110EV90W047zsU0y07zsFUSeWTkF09zXztDzHwV1U7zt00U14DodIm22Dn94EV90WEY8S0U0Cy4y9+2Y+aTkF0N8W49t4HwU484034EU1DEEmIE0281DwTz98WEY10UjwcW7zsFWQIWzwF088W498YG4Uk350V4EXl44KYYI0W81B0EV90WTY409849WLUdsG58WTkF08M449s4G0V009zV8GU100EW4E82817kUV90W0480284u3bzsw64b2C3zz/3U5s0sE7U00000000087UU0000000000000004") ;抗菌药停止时间 定位
Sleep,400

; MouseClick, left,  18,  888



MsgBox 术后首剂抗血小板药物使用时间
Sleep, 100
;术后首剂抗血小板药物使用时间
	FillDateTimeField("|<>*155$115.4800Fzw487zV1w024zk240U8V2zzm0FzWS1108DzvzYEV12108EF90W040802GDzV60U4DDYzxDm5zsO948F3wE244GE8Y92U434W49+2Dz3m9t4HwUTyAOFzwx140V14YW92E81298V24MW0EUyG94V87zV4YEV44F09zl90WTY20EW2EEbV8U498bUF825zsl188E04Ty88GE8U12U4EbU5sw68104s0sE7U0000000US0000000000E", {x1:10, y1:200, x2:1060, y2:900}, data.yyyy, data.mm, data.dd, 23, data.min) ;术后医嘱-是否使用抗血小板药物：首剂用药日期时间：
	send,{esc}
	; 术后72小时之后继续使用的原因：
Text:="|<>*158$152.8U083zw1zY4140204110U8E7zu8020W0821V0900U1+IFz3zt26byTz8U20U8E0E3zyWq92110EVfE402TwVz04040402ZUjyHz7zuX110YE8GL10zz3zb8Fl1AYF26U0Tz8824YEM0EFU8Hz42H94EVdz4A2zwV946044jy8X20UHz7zuGF50UU+GF2U2F20XdMs844F26YYGI8r2YYQY1WEzs250ztB0EVd994mkA996EUk4820t1lkEk88OTmV6U043l8AM320Vs0UW4T224Y40EDzt0U4187UVsXzUUFMQ0jU"


if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
	FindText().Click(X-96, Y+1, "L") ;临床医师认为有继续使用抗菌药物治疗的适应症的患者，
	send,{esc}
	
}

SendInput, {PGDN}
Sleep,300
rightsite("|<>*155$185.0803lzszz0U20Q08420401005040w8004Dw20E30103T1yE8Y28E40080E18E000E07zUCUzzU400UN84EVzszzbzWLw7zsU0813Yk80k807tjy8V2GE0U997d8EM1zyTy88EzsLzU2H8UTy4YVt4GG9GEVs200000030E0UTYeF024990GEYYGYU48403zyTy/zV108942244GG0YV98x90E8/z040U4412TsUm9zY88YYDC2GF/y30AI22Dl08Dy4UF94E88EF9828YYWIY40984AE20EE4F0Ym8UEEUWGE2l994V0E02TswU7zUU8XzAMF0Uzz4YXvGGG920U08UH7z8113l420HWzw02zzk8Pzzi400000000000000010000000000000001") ;;术后是否有活动性出血或血肿：
Sleep,400
;术后是否有活动性出血或血肿：
Text:="|<>*155$185.0803lzszz0U20Q08420401005040w8004Dw20E30103T1yE8Y28E40080E18E000E07zUCUzzU400UN84EVzszzbzWLw7zsU0813Yk80k807tjy8V2GE0U997d8EM1zyTy88EzsLzU2H8UTy4YVt4GG9GEVs200000030E0UTYeF024990GEYYGYU48403zyTy/zV108942244GG0YV98x90E8/z040U4412TsUm9zY88YYDC2GF/y30AI22Dl08Dy4UF94E88EF9828YYWIY40984AE20EE4F0Ym8UEEUWGE2l994V0E02TswU7zUU8XzAMF0Uzz4YXvGGG920U08UH7z8113l420HWzw02zzk8Pzzi400000000000000010000000000000001"

if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
		; FindText().Click(514, Y, "L") ;点击 是
  FindText().Click(600, Y, "L") ;点击 否
}
Sleep,30

;再手术：
Text:="|<>*165$46.TzU1k3M030Ts0Ak3zk300k08l0A3zz4X4TzUA0Hzk301s08l0A0/E0X4zzlAUDzw30MlU810A3334U40k0A0G3kS00k0U"


if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
		; FindText().Click(514, Y, "L") ;点击 是
  FindText().Click(600, Y, "L") ;点击 否
}
Sleep,30

;是否有手术后并发证：
Text:="|<>*152$130.Dz7zs40070100S42211Dz0U40k0E1zU013z08E912103zk7ETzk0000807zt40840813Yk80007zsU0247zw0E4zwEEVzlzw0k3zw8E00C90E0000A10006U800V0Dw8bUTznzlTw000V0U0zzkkEWE0108110Hzz422zk8E4W29017sU47z001U6+10V0FE8Y0AE20EE4004098444270mE5t0Dz10E20004zkUEFq290QTwU44D1s000W1410M23zkU"


if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
		; FindText().Click(514, Y, "L") ;点击 是
  FindText().Click(600, Y, "L") ;点击 否
}
Sleep,30

;是否进行术后头颅影像学检查：
Text:="|<>*157$186.Dz7zt280zU100S4E17xz24k28EUE000810A0W84000Ezk3E1l1144y14UUsTzUDz0R0jy80000U09E111z8947zvtA1k0813Yk28G07zsU0AE17t1E/zY08W6001Dz44/W85zkA0zz2E7o9z2N8VzUbxTzl00000W8A20S0U00E4Ic84NzU11kU810TznzkjzM20V0U0TzoI/zt8sU61cYDz00E20EW80210Ujw0E7o91G/DDzyWI8104TW0Ea80260Mc40s4J9z48Z040WIDz0AE20EY802409841a41UeA+AU40WM811SE3zlE002001Dw6142O9M8YU40U8001lzm0EDz0S00284M0UA4M0+40s0byzzk0000000000000000000000M00000000U"
if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
		; FindText().Click(514, Y, "L") ;点击 是
  FindText().Click(600, Y, "L") ;点击 否
}
Sleep,30





;出院带药医瞩：阿司匹林医瞩：
rightsite("|<>*158$185.10D40UE48Dznjw00Dzvzwzz223zwvz0W4K8DzzzzF05E800G0U094E444E1I2148fz210EUW0+zk00c100G8Vzz8U2jw28FI200161DyJ4101HmzyYF0kkHzZF0LzX7wzz4DmF0fz2034Y018W1lUYE+zkUV50000GUY41KG0059802F4XrV10JYUV1/zlzsx1/zmjw00+GFz4W9Ce2zwfz122EG20EFWEE5F000GYW29AGIZ441IE244cY40V14XQDzk00ht44GEx/88r3zw489l9817V9M6Ga101m2DsZ00IGK1YdUTzm4GEA00G005z2020AE1800UWU01TkU0Yks00D1bzw260041k0yTzl11zz0VU0000000US0000000000000000000001")
sleep,400
Text:="|<>*158$185.10D40UE48Dznjw00Dzvzwzz223zwvz0W4K8DzzzzF05E800G0U094E444E1I2148fz210EUW0+zk00c100G8Vzz8U2jw28FI200161DyJ4101HmzyYF0kkHzZF0LzX7wzz4DmF0fz2034Y018W1lUYE+zkUV50000GUY41KG0059802F4XrV10JYUV1/zlzsx1/zmjw00+GFz4W9Ce2zwfz122EG20EFWEE5F000GYW29AGIZ441IE244cY40V14XQDzk00ht44GEx/88r3zw489l9817V9M6Ga101m2DsZ00IGK1YdUTzm4GEA00G005z2020AE1800UWU01TkU0Yks00D1bzw260041k0yTzl11zz0VU0000000US0000000000000000000001"

if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
	
  FindText().Click(X+407, Y+51, "L") ;点击
}
SendInput, {DOWN}{ENTER}


;出院带药医瞩-β阻滞剂医瞩：
Text:="|<>*167$172.1UD40aE48Dzvjy0SDTnBg41bzvzy168gMTzzzyW0+UM3ApV3zs86F0/U84MWjw9Y12280fzUAHK43PDxNA0jzUFW+UE00Mk9zmcU0lCMGTy1ZZzmsEBzslzDzl3wgE+zk38tz509YKIE/jkkMWU01U9MG20f90BXq40M1lN10iZ0VX/zlzsx1/zmjwwnBMFDwspZzuvw26AgY4MVX4VU+W036pz4aEYKEk/V08MmmEFW46GBkzz0APK4GN2FN6szzUVXCN968w9/0mQo0lDMF9Y91ZUmAODzwX4YPU04U01Tk3slVAbl46E01zsk0mMQ1U7Unzy130A3DyUM8Flzy53U0000000US000000k000000000000000000000000000300000000000008"


if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
	
  FindText().Click(X+407, Y+51, "L") ;点击
}
SendInput, {DOWN}{ENTER}


;出院带药医嘱-他汀类药物医嘱：
Text:="|<>*164$191.1UD40aE48Dzvjy08M8zt4848283zyvzUX4KADzzzzF05EA0Ik8618XzzIE4E1I3168fz2N0EUW0+zs1dw8A1G0EUdz8U2jy2AFI200361DyJ402T/0MzzX61ueHzZF0LzX7wzz4Dml0fz0BaH0k3k4DodIgE+zkUl50030GkY41KG0tAW1UmQGkdGd10JYUVX/zlzsx1/zmjwwGN03644x139Gzwfz136KG2AElWEk5F00Yy86DzslWQIYA1IE26AgY4MV1YXQDzk19UEA0M11Zd98r3zw4ANn98l7V9M6HaU2H4UM187V8IWK1YtcTzmAGFg00G005z04UO0kQC00EW4U01TkU1Yks30D1bzw2608zYD3U6D1V9lzz0VU0000000US0000000000000US20000001"

if (ok:=FindText(X, Y, 130-150000, 574-150000, 130+150000, 574+150000, 0, 0, Text))
{
	
  FindText().Click(X+407, Y+51, "L") ;点击
}
SendInput, {DOWN}{ENTER}
;把特定题目移动到y=300
rightsite(Text){
    
; Text:= ;建立血管桥支数：
if (ok:=FindText(X, Y, 116-150000, 862-150000, 116+150000, 862+150000, 0, 0, Text))
{
	; FindText().Click(X, Y, "L")
	; MsgBox y=%y%
	move:=y-300
	; MsgBox move=%move%
	step:=Round(Abs(move/100), 1)
	; MsgBox step=%step%
	if move>0
		Send {WheelDown %step%}
	if move<0
		Send {WheelUp %step%}
}
}