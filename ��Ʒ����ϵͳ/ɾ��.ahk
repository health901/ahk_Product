#Include Func.ahk
database=%A_ScriptDir%\���ݿ�.mdb
Gui, Add, Text, x16 y20 w100 h30 , ����
Gui, Add, Edit, x126 y20 w180 h20 vID,
Gui, Add, Button, x316 y20 w90 h20 gDelete_item, ɾ����Ʒ
Gui, Add, Text, x16 y70 w100 h30 , ��ɫ
Gui, Add, Edit, x126 y70 w180 h20 vColor,
Gui, Add, Button, x316 y70 w90 h20 gDelete_Color, ɾ����ɫ
Gui, Add, Text, x16 y120 w100 h30 , Ʒ��
Gui, Add, Edit, x126 y120 w180 h20 vBrand,
Gui, Add, Button, x316 y120 w90 h20 gDelete_Brand, ɾ��Ʒ��
Gui, Add, Text, x16 y170 w390 h80 , ��������`,�����ɾ���ᵼ������ϵͳ�쳣
Gui +OwnDialogs
Gui, Show, h210 w416, ɾ��
Return

GuiClose:
ExitApp

Delete_item:
MsgBox,0x31,ɾ��ȷ��,ȷ��Ҫɾ������?
IfMsgBox,Ok
{
	GuiControlGet,ID
	Delete_ID(database,ID)
	}
IfMsgBox,Cancel
	Return
Return

Delete_Color:
MsgBox,0x31,ɾ��ȷ��,ȷ��Ҫɾ������ɫ?
IfMsgBox,Ok
{
	GuiControlGet,Color
	Delete_Color(database,Color)
	}
IfMsgBox,Cancel
	Return
Return

Delete_Brand:
MsgBox,0x31,ɾ��ȷ��,ȷ��Ҫɾ����Ʒ��?
IfMsgBox,Ok
{
	GuiControlGet,Brand
	Delete_Brand(database,Brand)
	}
IfMsgBox,Cancel
	Return
Return