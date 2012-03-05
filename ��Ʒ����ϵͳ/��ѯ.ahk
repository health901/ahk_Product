#Include Func.ahk
database=%A_ScriptDir%\���ݿ�.mdb
Menu, MyContextMenu, Add, �޸ĸ���Ʒ, Modify
Gui, Add, DropDownList, x16 y20 w100 h20 R10 vClass gChangeClass, ����||
Gui, Add, DropDownList, x136 y20 w100 h20 R10 vSubClass, ����(����)||
Gui, Add, DropDownList, x256 y20 w100 h20 R10 vQueryItem, ��ѯ��||SKU
Gui, Add, Edit, x386 y20 w470 h20 vDetail,
Gui, Add, Button, x886 y10 w90 h30 Default  gQuery, ��ѯ
Gui, Add, ListView, x16 y50 w680 h650 -Multi AltSubmit gMyLV vMyLV, ��Ʒ|���
OldSelect=
NewSelect=
LV_ModifyCol(1, 586),LV_ModifyCol(2, 90)
Gui, Add, Edit, x706 y180 w510 h520 ReadOnly vScreen,
Gui, Add, Picture, x716 y50 w120 h120 gLargePic vPIc, %A_ScriptDir%\��ƷͼƬ\default.png
Gui, Show, x122 y101 h708 w1226, ��Ʒ��ѯ
GetClass(database)
Return

ChangeClass:
GuiControlGet,Class,1:,Class
if class=
	Return
if Class=����
{
	GuiControl,1:,SubClass,|����(����)||
	GuiControl,1:,QueryItem,|��ѯ��||SKU
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
if Class=����
	Class=
if SubClass=����(����)
	SubClass=
if QueryItem=��ѯ��
	QueryItem=
if (QueryItem || SubClass || Class) =0
	Return
if (QueryItem || SubClass)=0		;����1����0��Ŀ0
	SQL:="Select ��Ʒ����,��� From " . Class
Else
	SQL:="Select ��Ʒ����,��� From " . Class . " Where "
if QueryItem=SKU
	Goto,Query_SKU
If SubClass!=
	SQL.="����='" . SubClass . "'"
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
	Return	;�ջ��ߴ���
GuiControl,,Screen,%Info%
if RegExMatch(info,"i)(?<=:\s)\\.*?\.(jpg|jpeg|png|gif)",Picture)
	Pic=%A_Scriptdir%%Picture%
Else
	Pic=%A_Scriptdir%\��ƷͼƬ\default.png
GuiControl,,Pic,%Pic%
Return

GuiClose:
ExitApp

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;�Ҽ��˵�
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
	IfNotExist,�޸�.ahk  ;exe
		Return
;~ 	Run,�޸�.exe %id%
	Run,`"%A_AhkPath%`" �޸�.ahk %ID%
	}
Return