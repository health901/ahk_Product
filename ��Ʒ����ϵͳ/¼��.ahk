#Include Func.ahk
database=%A_ScriptDir%\���ݿ�.mdb
Gui, Add, Picture, x16 y13 w200 h200 vItem_Pic, %A_ScriptDir%\��ƷͼƬ\Default.png
Gui, Add, GroupBox, x6 y0 w220 h220 ,
Gui, Font ,s20 bold,����
Gui, Add, Text, x296 y10 w800 h40 cred vTip,
Gui, Font
Gui, Add, Text, x236 y60 w40 h20 , ����
Gui, Add, Edit, x286 y60 w615 h20 vItem,
Gui, Add, Text, x236 y110 w40 h20 , ͼƬ
Gui, Add, Edit, x286 y110 w570 h20 ReadOnly vItem_PicPath,
Gui, Add, Button, x866 y110 w35 h20 gSelectPic, ...
Gui, Add, Text, x236 y170 w40 h20 , ����
Gui, Add, DropDownList, x286 y170 w70 h20 R10 vClass gChangeClass,
GetClass(database)
Gui, Show, x134 y88 h228 w1183, ¼��ϵͳ
Gui,-MaximizeBox -MinimizeBox
WinGet,hwnd,id,¼��ϵͳ
OnMessage(0x47,"WindowsMove")
WinActivate,¼��ϵͳ
Gui2Exist=0
Return

GuiClose:
ExitApp

GUI2show:
Gui2Exist=1
Gui, 2:destroy
Gui, 2:+owner1
if	Class=ң�س�
	Gosub,Gui2_ykc			;ң�س���
Else if Class=��������
	Gosub,Gui2_zn
Else
	return
GetColor(database)
WinGetPos,X,Y,,,ahk_id %hwnd%
Y+=263
gui,2:show,x%X% y%Y%  h386 w1183, ��ϸ��Ϣ¼��
gui,2:-SysMenu
Return

SelectPic:
FileSelectFile,Item_Pic,1,,,ͼƬ(*.jpg;*.png;*.gif;*.jpeg;*.tif)
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
		GuiControl,2:,Brand,|Ʒ��
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
	Item_Pic=\��ƷͼƬ\%ID%.%Ext%
	FileMove,%Item_Pic_O%,%A_ScriptDir%%Item_Pic%,1
}
if	Class=ң�س�
	Gosub,Write2DB_ykc
Else if Class=��������
	Gosub,Write2DB_zn
Else
	return
SetTimer,TipClear,3000
return


Write2DB_ykc:
GuiControlGet,Class,1:,Class
GuiControlGet,Item,1:,Item
if brand=Ʒ��
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
if brand=Ʒ��
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
	;���ݴ���
	if (PackSize_length && PackSize_Width && PackSize_Width) = 0
		PackSize=
	Else
		PackSize:=PackSize_length . "*" PackSize_Width . "*" . PackSize_Width
	if (ItemSize_length && ItemSize_Width && ItemSize_Width) = 0
		ItemSize=
	Else
		ItemSize:=ItemSize_length . "*" ItemSize_Width . "*" . ItemSize_Height
	;����ɫ/��ɫ
	StringSplit,Color_,Color,/
	ColorCode=
	Loop,% Color_0
	{
		tmp:=GetColorCode(Database,Color_%A_index%)
		ColorCode.=tmp "/"
	}
	tmp=
	StringTrimRight,ColorCode,ColorCode,1
	;д�����ݿ�
	com_init() ; ��ʼ��COM
	pcon := COM_CreateObject("ADODB.Connection") ; ��������
	com_invoke(pcon, "open", "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" . database)     ; ������
	SQL:="INSERT INTO " . Class . "(���,Ʒ��,�ͺ�,��Ʒ����,��ģ����,��װ�ߴ�,��Ʒ�ߴ�,����,��Ʒ����,��ɫ,�Ƿ��й�,��������,��Ʒ��װ,�۸�,��ע,ͼƬ,����,����,����) values('" . ID . "','" . Brand . "','" . Model . "','" . Item . "','" . Scale . "','" . PackSize . "','" . ItemSize . "','" . Material . "','" . Weight . "','" . ColorCode . "','" . Light . "','" . Age . "','" . Package . "','" . Price . "','" . PS . "','" . Item_Pic . "','" . PurchaseCode . "','" . Origin . "','" . Inventory . "')"
	com_invoke(pcon, "Execute", SQL)  ; ִ�� SQL���
	com_invoke(pcon, "close")     ; �ر�����
	com_term()
	GuiControl,1:,Tip,¼��ɹ�
Return
Write_zn:
	;���ݴ���
	if (PackSize_length && PackSize_Width && PackSize_Width) = 0
		PackSize=
	Else
		PackSize:=PackSize_length . "*" PackSize_Width . "*" . PackSize_Width
	if (ItemSize_length && ItemSize_Width && ItemSize_Width) = 0
		ItemSize=
	Else
		ItemSize:=ItemSize_length . "*" ItemSize_Width . "*" . ItemSize_Height
	;����ɫ/��ɫ
	StringSplit,Color_,Color,/
	ColorCode=
	Loop,% Color_0
	{
		tmp:=GetColorCode(Database,Color_%A_index%)
		ColorCode.=tmp "/"
	}
	tmp=
	StringTrimRight,ColorCode,ColorCode,1
	;д�����ݿ�
	com_init() ; ��ʼ��COM
	pcon := COM_CreateObject("ADODB.Connection") ; ��������
	com_invoke(pcon, "open", "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" . database)     ; ������
	SQL:="INSERT INTO " . Class . "(���,Ʒ��,�ͺ�,��Ʒ����,����,��װ�ߴ�,��Ʒ�ߴ�,����,��Ʒ����,��ɫ,��������,��Ʒ��װ,�۸�,��ע,ͼƬ,����,����) values('" . ID . "','" . Brand . "','" . Model . "','" . Item . "','" . Ability . "','" . PackSize . "','" . ItemSize . "','" . Material . "','" . Weight . "','" . ColorCode . "','" . Age . "','" . Package . "','" . Price . "','" . PS . "','" . Item_Pic . "','" . PurchaseCode . "','" . Origin .  "')"
	com_invoke(pcon, "Execute", SQL)  ; ִ�� SQL���
	com_invoke(pcon, "close")     ; �ر�����
	com_term()
	GuiControl,1:,Tip,¼��ɹ�
Return


ErrorInfo:
GuiControl,1:,Tip,����`,����`,Ʒ��`,�ͺ�`,�۸�`,���� Ϊ������,����
SetTimer,TipClear,3000
Return

TipClear:
SetTimer,TipClear,off
GuiControl,1:,Tip,
Return

Clear:
Gui,2:destroy
Gosub,GUI2show
WinActivate,��ϸ��Ϣ¼��
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
	gui,2:show,x%X% y%Y%  h386 w1183, ��ϸ��Ϣ¼��
	Return
}


;����GUI
#Include Gui2_ykc.ahk
#Include Gui2_zn.ahk