GuiLogin:
Gui, Add, Text, x16 y20 w100 h30 , 用户名:
Gui, Add, Edit, x126 y20 w130 h20 , Edit
Gui, Add, Text, x16 y70 w100 h30 , 密码:
Gui, Add, Edit, x126 y70 w130 h20 , Edit
Gui, Add, Button, x6 y130 w100 h30 , Login
Gui, Add, Button, x156 y130 w100 h30 , Exit
Gui, Show, x131 y91 h194 w278, 产品管理 - 登陆
Return



GuiBoard:
Gui, Add, Button, x26 y20 w110 h40 , 产品录入
Gui, Add, Button, x176 y20 w110 h40 , 产品查询
Gui, Add, Button, x26 y100 w110 h40 , 产品修改
Gui, Add, Button, x176 y100 w110 h40 , 删除项目
Gui, Show, x334 y179 h168 w321, 控制面板
Return

GuiClose:
ExitApp