#Include Func.ahk
database=%A_ScriptDir%\数据库.mdb
Gui, Add, Text, x16 y20 w100 h30 , 编码
Gui, Add, Edit, x126 y20 w180 h20 vID,
Gui, Add, Button, x316 y20 w90 h20 gDelete_item, 删除物品
Gui, Add, Text, x16 y70 w100 h30 , 颜色
Gui, Add, Edit, x126 y70 w180 h20 vColor,
Gui, Add, Button, x316 y70 w90 h20 gDelete_Color, 删除颜色
Gui, Add, Text, x16 y120 w100 h30 , 品牌
Gui, Add, Edit, x126 y120 w180 h20 vBrand,
Gui, Add, Button, x316 y120 w90 h20 gDelete_Brand, 删除品牌
Gui, Add, Text, x16 y170 w390 h80 , 谨慎操作`,错误的删除会导致整个系统异常
Gui +OwnDialogs
Gui, Show, h210 w416, 删除
Return

GuiClose:
ExitApp

Delete_item:
MsgBox,0x31,删除确认,确认要删除该项?
IfMsgBox,Ok
{
	GuiControlGet,ID
	Delete_ID(database,ID)
	}
IfMsgBox,Cancel
	Return
Return

Delete_Color:
MsgBox,0x31,删除确认,确认要删除该颜色?
IfMsgBox,Ok
{
	GuiControlGet,Color
	Delete_Color(database,Color)
	}
IfMsgBox,Cancel
	Return
Return

Delete_Brand:
MsgBox,0x31,删除确认,确认要删除该品牌?
IfMsgBox,Ok
{
	GuiControlGet,Brand
	Delete_Brand(database,Brand)
	}
IfMsgBox,Cancel
	Return
Return