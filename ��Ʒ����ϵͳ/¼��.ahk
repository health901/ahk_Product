#Include Func.ahk
database=%A_ScriptDir%\数据库.mdb
Gui, Add, Picture, x16 y13 w200 h200 vItem_Pic, %A_ScriptDir%\产品图片\Default.png
Gui, Add, GroupBox, x6 y0 w220 h220 ,
Gui, Font ,s20 bold,宋体
Gui, Add, Text, x296 y10 w800 h40 cred vTip,
Gui, Font
Gui, Add, Text, x236 y60 w40 h20 , 名称
Gui, Add, Edit, x286 y60 w615 h20 vItem,
Gui, Add, Text, x236 y110 w40 h20 , 图片
Gui, Add, Edit, x286 y110 w570 h20 ReadOnly vItem_PicPath,
Gui, Add, Button, x866 y110 w35 h20 gSelectPic, ...
Gui, Add, Text, x236 y170 w40 h20 , 分类
Gui, Add, DropDownList, x286 y170 w70 h20 R10 vClass gChangeClass,
GetClass(database)
Gui, Show, x134 y88 h228 w1183, 录入系统
Gui,-MaximizeBox -MinimizeBox
WinGet,hwnd,id,录入系统
OnMessage(0x47,"WindowsMove")
WinActivate,录入系统
Gui2Exist=0
Return

GuiClose:
ExitApp

GUI2show:
Gui2Exist=1
Gui, 2:destroy
Gui, 2:+owner1
if	Class=遥控车
	Gosub,Gui2_ykc			;遥控车段
Else if Class=智能娃娃
	Gosub,Gui2_zn
Else
	return
GetColor(database)
WinGetPos,X,Y,,,ahk_id %hwnd%
Y+=263
gui,2:show,x%X% y%Y%  h386 w1183, 详细信息录入
gui,2:-SysMenu
Return

SelectPic:
FileSelectFile,Item_Pic,1,,,图片(*.jpg;*.png;*.gif;*.jpeg;*.tif)
if item_Pic=
	Return
GuiControl,1:,Item_Pic,%Item_Pic%
GuiControl,1:,Item_PicPath,%Item_Pic%
Return


ChangeClass:
GuiControlGet,Class,1:,Class
if class=
{
	if Gui2Exist=0
	{
		GuiControl,2:,Brand,|品牌
		Return
		}
}
Gosub,GUI2Show
GetBrand(database,class)
CreatID(database,Class)
Return

AddBrand:
GuiControlGet,Class,1:,Class
if class=
	Return
GuiControlGet,NewBrand,2:,NewBrand
if Newbrand=
	Return
ADDBrand(database,NewBrand,class)
GetBrand(database,Class)
Return

AddColor:
GuiControlGet,newColor,2:,newColor
if newcolor=
	Return
AddColor(database,newColor)
GetColor(database)
Return

Write2DB:
if Item_Pic!=
{
	GuiControlGet,ID,2:,ID
	SplitPath,Item_Pic,,,Ext
	Item_Pic_O:=Item_Pic
	Item_Pic=\产品图片\%ID%.%Ext%
	FileMove,%Item_Pic_O%,%A_ScriptDir%%Item_Pic%,1
}
if	Class=遥控车
	Gosub,Write2DB_ykc
Else if Class=智能娃娃
	Gosub,Write2DB_zn
Else
	return
SetTimer,TipClear,3000
return


Write2DB_ykc:
GuiControlGet,Class,1:,Class
GuiControlGet,Item,1:,Item
if brand=品牌
	brand=0
Color=
RowNumber = 0
Loop
{
    RowNumber := LV_GetNext(RowNumber,"Checked")
    if not RowNumber  ;
        break
    LV_GetText(Text, RowNumber)
    Color.=Text . "/"
}
StringTrimRight,color,color,1
gui,2:submit,nohide
if (brand && Model && Price && Color && Class && Item && PurchaseCode) =0
	Gosub,ErrorInfo
Else
	Gosub,Write_ykc
Return

Write2DB_zn:
GuiControlGet,Class,1:,Class
GuiControlGet,Item,1:,Item
if brand=品牌
	brand=0
Color=
RowNumber = 0
Loop
{
    RowNumber := LV_GetNext(RowNumber,"Checked")
    if not RowNumber  ;
        break
    LV_GetText(Text, RowNumber)
    Color.=Text . "/"
}
StringTrimRight,color,color,1
gui,2:submit,nohide
if (brand && Model && Price && Color && Class && Item && PurchaseCode) =0
	Gosub,ErrorInfo
Else
	Gosub,Write_zn
Return

Write_ykc:
	;数据处理
	if (PackSize_length && PackSize_Width && PackSize_Width) = 0
		PackSize=
	Else
		PackSize:=PackSize_length . "*" PackSize_Width . "*" . PackSize_Width
	if (ItemSize_length && ItemSize_Width && ItemSize_Width) = 0
		ItemSize=
	Else
		ItemSize:=ItemSize_length . "*" ItemSize_Width . "*" . ItemSize_Height
	;咖啡色/红色
	StringSplit,Color_,Color,/
	ColorCode=
	Loop,% Color_0
	{
		tmp:=GetColorCode(Database,Color_%A_index%)
		ColorCode.=tmp "/"
	}
	tmp=
	StringTrimRight,ColorCode,ColorCode,1
	;写入数据库
	com_init() ; 初始化COM
	pcon := COM_CreateObject("ADODB.Connection") ; 创建连接
	com_invoke(pcon, "open", "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" . database)     ; 打开连接
	SQL:="INSERT INTO " . Class . "(编号,品牌,型号,产品名称,车模比例,包装尺寸,产品尺寸,材质,产品重量,颜色,是否有光,适用年龄,产品包装,价格,备注,图片,货号,产地,配置) values('" . ID . "','" . Brand . "','" . Model . "','" . Item . "','" . Scale . "','" . PackSize . "','" . ItemSize . "','" . Material . "','" . Weight . "','" . ColorCode . "','" . Light . "','" . Age . "','" . Package . "','" . Price . "','" . PS . "','" . Item_Pic . "','" . PurchaseCode . "','" . Origin . "','" . Inventory . "')"
	com_invoke(pcon, "Execute", SQL)  ; 执行 SQL语句
	com_invoke(pcon, "close")     ; 关闭连接
	com_term()
	GuiControl,1:,Tip,录入成功
Return
Write_zn:
	;数据处理
	if (PackSize_length && PackSize_Width && PackSize_Width) = 0
		PackSize=
	Else
		PackSize:=PackSize_length . "*" PackSize_Width . "*" . PackSize_Width
	if (ItemSize_length && ItemSize_Width && ItemSize_Width) = 0
		ItemSize=
	Else
		ItemSize:=ItemSize_length . "*" ItemSize_Width . "*" . ItemSize_Height
	;咖啡色/红色
	StringSplit,Color_,Color,/
	ColorCode=
	Loop,% Color_0
	{
		tmp:=GetColorCode(Database,Color_%A_index%)
		ColorCode.=tmp "/"
	}
	tmp=
	StringTrimRight,ColorCode,ColorCode,1
	;写入数据库
	com_init() ; 初始化COM
	pcon := COM_CreateObject("ADODB.Connection") ; 创建连接
	com_invoke(pcon, "open", "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" . database)     ; 打开连接
	SQL:="INSERT INTO " . Class . "(编号,品牌,型号,产品名称,功能,包装尺寸,产品尺寸,材质,产品重量,颜色,适用年龄,产品包装,价格,备注,图片,货号,产地) values('" . ID . "','" . Brand . "','" . Model . "','" . Item . "','" . Ability . "','" . PackSize . "','" . ItemSize . "','" . Material . "','" . Weight . "','" . ColorCode . "','" . Age . "','" . Package . "','" . Price . "','" . PS . "','" . Item_Pic . "','" . PurchaseCode . "','" . Origin .  "')"
	com_invoke(pcon, "Execute", SQL)  ; 执行 SQL语句
	com_invoke(pcon, "close")     ; 关闭连接
	com_term()
	GuiControl,1:,Tip,录入成功
Return


ErrorInfo:
GuiControl,1:,Tip,名称`,类型`,品牌`,型号`,价格`,货号 为必填项,请检查
SetTimer,TipClear,3000
Return

TipClear:
SetTimer,TipClear,off
GuiControl,1:,Tip,
Return

Clear:
Gui,2:destroy
Gosub,GUI2show
WinActivate,详细信息录入
GetBrand(Database,Class)
Return


WindowsMove(wParam, lParam, msg, hwnd1)
{
	global hwnd
	if hwnd1<>%hwnd%
	Return
	GuiControlGet,Class,1:,Class
	if class=
		Return
	WinGetPos,X,Y,,,ahk_id %hwnd%
	Y+=263
	gui,2:show,x%X% y%Y%  h386 w1183, 详细信息录入
	Return
}


;分类GUI
#Include Gui2_ykc.ahk
#Include Gui2_zn.ahk