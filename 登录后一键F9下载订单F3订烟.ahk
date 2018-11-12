SetWorkingDir %A_ScriptDir%
global gapTime:=100
global IDorder := Array()
order := getOrderFromCSV()

F1::
getOrderFromCSV()
Return
F2::
Loop, read, 订烟表.csv
{
    Loop, parse, A_LoopReadLine, CSV
    {
        if (A_Index==7&&A_LoopField!="")
        {
            value:=A_LoopField
            value+=0
            gapTime:=value
        }
    }
}
MsgBox, %gapTime%
Return
F4::
for k,v in IDorder
    {
       MsgBox,%k%,%v%
    }
     
Return
;订烟
F3::

;构建订单kv表
; order := Array()
; order["椰树(软)"]:=5
; order["羊城(软红)"]:=5
; order["广州双喜(软盛世)"]:=2
; order["广州双喜(硬世纪经典)"]:=1
; order["广州双喜(硬经典1906)"]:=10
; order["广州双喜(硬百年经典)"]:=1
; order["双喜(软珍品好日子)"]:=5
; order["双喜(硬金樽好日子)"]:=10
; order["双喜(硬精品好日子)"]:=20
; order["红玫(硬金)"]:=2
; order["牡丹(软)"]:=4
; order["中华(硬)"]:=15
; order["玉溪(软)"]:=5
; order["玉溪(软尚善)"]:=5
; order["红塔山(硬经典)"]:=10
; order["黄金叶(硬金满堂)"]:=5
; order["白沙(软)"]:=5
; order["芙蓉王(硬)"]:=35
; order["白沙(硬精品)"]:=10
; order["南京(红)"]:=3
; order["南京(硬炫赫门)"]:=3
; order["黄鹤楼(硬天下名楼)"]:=1
; order["利群(硬新版)"]:=5


;进入订烟界面
; pwb := IEGet("首页 - 新商盟")
; if !pwb
; pwb := IEGet("新商盟")
; doc := pwb.Document

;开始订烟
pwb := IEGet("首页 - 新商盟")
if !pwb
pwb := IEGet("新商盟")
doc := pwb.Document

idorderlength := 0
for k,v in IDorder
{
    idorderlength++
    if (idorderlength>1)
    {
        Goto ,HASID
    }
}
;MsgBox,"-1"
NOID:
{
    
i:=doc.all.Length
index:=0
;MsgBox,"0"
Loop %i%
{   
    ;MsgBox,"1"
    counter := 0
    for k, v in order
        counter++
    if (counter==0)
    {
        doc.all[index+6-23].focus()
        Sleep,gapTime
        Break
    }
    element := doc.all[A_Index-1]
    if element.tagName != "span"
        Continue
    index:=A_Index-1
      ;MsgBox,"2"
    for k,v in order
    {
        ;MsgBox,%k%,%v%
        if (k=element.innerText)
        {
           
            tar:=doc.all[index+6]
            tar.focus()
            tar.value:=v
            IDorder[index+6]:=v
            order.Delete(k)
            Sleep,gapTime
            Continue
        }
    }
}
Goto, ENDed
 ;MsgBox,"end"
}
HASID:
{
    for k,v in IDorder
    {
        tar:=doc.all[k]
        tar.focus()

            tar.value:=v

            Sleep,gapTime
            index:=k-6
    }
     doc.all[index+6-23].focus()
        Sleep,gapTime
}
ENDed:
Sleep,gapTime
pwb.document.getElementById("smt").Click()

Return



;functions
IEGet(Name="")        ;Retrieve pointer to existing IE window/tab
{
    thewb := 0
    IfEqual, Name,, WinGetTitle, Name, ahk_class IEFrame
        Name := ( Name="New Tab - Windows Internet Explorer" ) ? "about:Tabs"
        : RegExReplace( Name, " - (Windows|Microsoft) Internet Explorer" )
    For wb in ComObjCreate( "Shell.Application" ).Windows
        If ( wb.LocationName = Name ) && InStr( wb.FullName, "iexplore.exe" )
            thewb:=wb
    If (thewb!=0)
        Return thewb
            
} ;written by Jethrow
; wb := IEGet() ;Last active window/tab
; ;OR
; wb := IEGet("Google") ;Tab name you define can also be a variable
IELoad(wb)    ;You need to send the IE handle to the function unless you define it as global.
{
    If !wb    ;If wb is not a valid pointer then quit
        Return False
    Loop    ;Otherwise sleep for .1 seconds untill the page starts loading
        Sleep,100
    Until (wb.busy)
    Loop    ;Once it starts loading wait until completes
        Sleep,100
    Until (!wb.busy)
    Loop    ;optional check to wait for the page to completely load
        Sleep,100
    Until (wb.Document.Readystate == "complete")
Return True
}
getOrderFromCSV()
{
order:=Array()
Loop, read, 订烟表.csv
{
    key:=-1
    value:=-1
    Loop, parse, A_LoopReadLine, CSV
    {
        if (A_Index==1&&A_LoopField!="")
        {
            key:=A_LoopField
        }
        if (A_Index==2&&A_LoopField!="")
        {
            value:=A_LoopField
        }
        if (A_Index==7&&A_LoopField!="")
        {
            value:=A_LoopField
            value+=0
            gapTime:=value
            
        }
    }
    if (key!=-1 && value!=-1)
    {
        order[key]:=value
    }
}

Return order
}