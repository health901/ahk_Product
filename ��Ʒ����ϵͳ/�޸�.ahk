#Include Func.ahk
database=%A_ScriptDir%\数据库.mdb
#SingleInstance,force		;新实例替换旧实例
Gui, Add, Picture, x16 y13 w200 h200 vItem_Pic, %A_ScriptDir%\产品图片\Default.png
Gui, Add, GroupBox, x6 y0 w220 h220 ,
Gui, Font, s20 bold, 宋体
Gui, Add, Text, x296 y10 w800 h40 cred vTip,
Gui, Font
Gui, Add, Text, x236 y60 w40 h20 , 名称
Gui, Add, Edit, x286 y60 w615 h20 vItem,
Gui, Add, Text, x236 y110 w40 h20 , 图片
Gui, Add, Edit, x286 y110 w570 h20 ReadOnly vItem_PicPath,
Gui, Add, Button, x866 y110 w35 h20 gSelectPic, ...
Gui, Add, Text, x236 y153 h20 , 编号/SKU:
Gui, Add, Edit, x297 y150 w100 h20 vID_G1,
Gui, Add, Button, x422 y150 w100 h30 gLoadInfo, 载入
Gui, Add, Button, x572 y150 w100 h30 gSave, 保存修改
Gui, Show, h232 w1187, 信息修改
if 0<>0
{
	GuiControl,,ID_G1,%1%
	}
Return

GuiClose:
ExitApp


SelectPic:
FileSelectFile,Item_Pic,1,,,图片(*.jpg;*.png;*.gif;*.jpeg;*.tif)
if item_Pic=
	Return
GuiControl,1:,Item_Pic,%Item_Pic%
GuiControl,1:,Item_PicPath,%Item_Pic%
Return

Loadinfo:
GuiControlGet,ID_G1
info:=GetAllInfo(database,ID)
if RegExMatch(info,"i)(?<=:\s)\\.*?\.(jpg|jpeg|png|gif)",Picture)
	Pic=%A_Scriptdir%%Picture%
Else
	Pic=%A_Scriptdir%\产品图片\default.png
GuiControl,,Item_Pic,%Pic%
pos:=RegExMatch(ID_G1,"[[:alpha:]]*?(?=\d)",ClassCode)
if !pos
	Return
Gosub,load_%Classcode%
Loop,parse,info,`n,`r
{
	Pos:=RegExMatch(A_LoopField,"(?<=:\s).*",field)
	tmp:=Var_%a_index%
	%tmp%:=field
}
Return

Save:
Return

;;;;;;;;;;;;;;;;;;;;;;;;;;
;分类load

Load_ykc:
VarCount=19
Var_1=ID
Var_2=Brand
Var_3=Model
Var_4=Item
Var_5=Scale
Var_6=PackSize
Var_7=ItemSize
Var_8=Material
Var_9=Weight
Var_10=ColorCode
Var_11=Light
Var_12=Age
Var_13=Package
Var_14=Price
Var_15=PS
Var_16=Item_Pic
Var_17=PurchaseCode
Var_18=Origin
Var_19=Inventory
Loop,19
{
	tmp:=Var_%a_index%
	%tmp%=
	}
Return

Load_zn:
VarCount=19
Var_1=ID
Var_2=Brand
Var_3=Model
Var_4=Item
Var_5=Ability
Var_6=PackSize
Var_7=ItemSize
Var_8=Material
Var_9=Weight
Var_10=ColorCode
Var_11=Age
Var_12=Package
Var_13=Price
Var_14=PS
Var_15=Item_Pic
Var_16=PurchaseCode
Var_17=Origin
Loop,17
{
	tmp:=Var_%a_index%
	%tmp%=
	}
Return