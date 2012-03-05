Gui2_ykc:
Gui, 2:Add, DropDownList, x6 y10 w80 h20 R10 vBrand, 品牌||
Gui, 2:Add, Button, x96 y10 w60 h20 gAddBrand, 添加
Gui, 2:Add, Edit, x196 y10 w70 h20 vNewBrand, 新品牌名
Gui, 2:Add, Text, x6 y50 w70 h20 , 型号
Gui, 2:Add, Edit, x86 y50 w100 h20 vModel,
Gui, 2:Add, Text, x6 y90 w70 h20 , 模型比例
Gui, 2:Add, Edit, x86 y90 w100 h20 vScale,
Gui, 2:Add, Text, x6 y130 w70 h20 , 材质
Gui, 2:Add, Edit, x86 y130 w100 h20 vMaterial,
Gui, 2:Add, Text, x6 y170 w70 h20 , 产品重量
Gui, 2:Add, Edit, x86 y170 w100 h20 vWeight,
Gui, 2:Add, Text, x6 y210 w70 h20 , 是否有光
Gui, 2:Add, Edit, x86 y210 w100 h20 vLight,
Gui, 2:Add, Text, x6 y250 w70 h20 , 适用年龄
Gui, 2:Add, Edit, x86 y250 w100 h20 vAge, 3岁以上
Gui, 2:Add, Text, x6 y290 w70 h20 , 产品包装
Gui, 2:Add, Edit, x86 y290 w100 h20 vPackage, 开窗彩盒
Gui, 2:Add, Text, x6 y340 w70 h20 , 价格
Gui, 2:Add, Edit, x86 y340 w100 h20 vPrice,
Gui, 2:Add, Text, x196 y40 w70 h20 , 包装尺寸
Gui, 2:Add, Text, x196 y70 w40 h20 , 长
Gui, 2:Add, Text, x256 y70 w40 h20 , 宽
Gui, 2:Add, Text, x306 y70 w40 h20 , 高
Gui, 2:Add, Edit, x196 y100 w40 h20 vPackSize_length,
Gui, 2:Add, Edit, x256 y100 w40 h20 vPackSize_width,
Gui, 2:Add, Edit, x306 y100 w40 h20 vPackSize_height,
Gui, 2:Add, Text, x196 y140 w70 h20 , 产品尺寸
Gui, 2:Add, Text, x196 y170 w40 h20 , 长
Gui, 2:Add, Text, x256 y170 w40 h20 , 宽
Gui, 2:Add, Text, x306 y170 w40 h20 , 高
Gui, 2:Add, Edit, x196 y200 w40 h20 vitem_length,
Gui, 2:Add, Edit, x256 y200 w40 h20 vitem_widtht,
Gui, 2:Add, Edit, x306 y200 w40 h20 vitem_height,
Gui, 2:Add, Text, x196 y263 w40 h20 , 产地
Gui, 2:Add, Edit, x236 y260 w110 h20 vOrigin,
Gui, 2:Add, Text, x196 y290 w130 h20 , 配置
Gui, 2:Add, Edit, x196 y320 w150 h40 vInventory,
Gui, 2:Add, ListView, x356 y10 w120 h300 Checked vColorList, 颜色列表
Gui, 2:Add, Edit, x356 y315 w120 h20 vNewColor,添加新颜色
Gui, 2:Add, Button, x356 y340 w120 h20 gAddColor, 添加
Gui, 2:Add, Text, x486 y10 w130 h20 , 备注
Gui, 2:Add, Edit, x486 y40 w180 h320 vPS,
Gui, 2:Add, Text, x696 y10 w110 h30 , 自动生成编码
Gui, 2:Add, Edit, x696 y50 w110 h50 ReadOnly vID,
Gui, 2:Add, Text, x696 y120 w110 h60 , 一种商品一个编号`n`n相同颜色编号相同`n`nSKU=货号+颜色码
Gui, 2:Add, Text, x696 y195 w110 h20 , 货号(进货编号)
Gui, 2:Add, Edit, x696 y215 w110 h20 vPurchaseCode,
Gui, 2:Add, Button, x696 y240 w110 h30 gWrite2DB, 录入数据库
Gui, 2:Add, Button, x696 y290 w110 h30 gClear, 清空
Return