#Include Func.ahk
database=%A_ScriptDir%\数据库.mdb
Menu, MyContextMenu, Add, 修改该物品, Modify
Gui, Add, DropDownList, x16 y20 w100 h20 R10 vClass gChangeClass, 大类||
Gui, Add, DropDownList, x136 y20 w100 h20 R10 vSubClass, 子类(暂无)||
Gui, Add, DropDownList, x256 y20 w100 h20 R10 vQueryItem, 查询项||SKU
Gui, Add, Edit, x386 y20 w470 h20 vDetail,
Gui, Add, Button, x886 y10 w90 h30 Default  gQuery, 查询
Gui, Add, ListView, x16 y50 w680 h650 -Multi AltSubmit gMyLV vMyLV, 商品|编号
OldSelect=
NewSelect=
LV_ModifyCol(1, 586),LV_ModifyCol(2, 90)
Gui, Add, Edit, x706 y180 w510 h520 ReadOnly vScreen,
Gui, Add, Picture, x716 y50 w120 h120 gLargePic vPIc, %A_ScriptDir%\产品图片\default.png
Gui, Show, x122 y101 h708 w1226, 产品查询
GetClass(database)
Return

ChangeClass:
GuiControlGet,Class,1:,Class
if class=
	Return
if Class=大类
{
	GuiControl,1:,SubClass,|子类(暂无)||
	GuiControl,1:,QueryItem,|查询项||SKU
	Return
}
GetQueryItem(database,Class)
Gosub,GetSubClass
Return

GetSubClass:
Return

LargePic:
Run,% Pic
Return

Query:
Gui,Submit,nohide
LV_Delete()
if Class=大类
	Class=
if SubClass=子类(暂无)
	SubClass=
if QueryItem=查询项
	QueryItem=
if (QueryItem || SubClass || Class) =0
	Return
if (QueryItem || SubClass)=0		;大类1子类0项目0
	SQL:="Select 产品名称,编号 From " . Class
Else
	SQL:="Select 产品名称,编号 From " . Class . " Where "
if QueryItem=SKU
	Goto,Query_SKU
If SubClass!=
	SQL.="子类='" . SubClass . "'"
if QueryItem!=
{
	IF DETAIL=
		RETURN
	SQL.=QueryItem . " like '%" . Detail . "%'"
}
LV_Delete()
Query(database,sql)
Return

Query_SKU:
	QuerySku(database,detail)
Return

MyLV:
Critical
RowNumber := LV_GetNext(0,"F")
LV_GetText(Text, RowNumber,2)
OldSelect:=NewSelect
NewSelect:=Text
if OldSelect=%NewSelect%
	Return
info:=GetAllInfo(database,Text)
if info=-1
	Return	;空或者错误
GuiControl,,Screen,%Info%
if RegExMatch(info,"i)(?<=:\s)\\.*?\.(jpg|jpeg|png|gif)",Picture)
	Pic=%A_Scriptdir%%Picture%
Else
	Pic=%A_Scriptdir%\产品图片\default.png
GuiControl,,Pic,%Pic%
Return

GuiClose:
ExitApp

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;右键菜单
GuiContextMenu:
if A_GuiControl <> MyLV
	Return
Menu, MyContextMenu, Show, %A_GuiX%, %A_GuiY%
Return

Modify:
FocusedRowNumber := LV_GetNext(0)
if not FocusedRowNumber
    return
LV_GetText(ID, FocusedRowNumber, 2)
if ID!=
{
	IfNotExist,修改.ahk  ;exe
		Return
;~ 	Run,修改.exe %id%
	Run,`"%A_AhkPath%`" 修改.ahk %ID%
	}
Return