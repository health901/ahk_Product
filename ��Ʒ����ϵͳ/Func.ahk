GetClass(database)
{
	COM_Init() ; 初始化COM
	prs := COM_CreateObject("ADODB.Recordset") ; 创建数据集 prs
	sSource  := "SELECT 类型 from 产品" ; SQL查询语句
	com_invoke(prs, "open", sSource, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" . database) ; 打开数据库
	Loop { ; 查询
		if com_invoke(prs, "EOF") ; 当碰到EOF标志时，退出循环
			Break
		Class := com_invoke(prs, "Fields[0].Value")    ; 获取数据集中第1个字段的值
		GuiControl,1:,Class,%Class%
		com_invoke(prs, "MoveNext")  ; 移到下一条
	}
	com_invoke(prs, "close") ; 关闭数据集
	com_release(prs)         ; 释放数据集
	COM_Term()
	Return
}
GetBrand(database,Class)
{
	if Class=
		Return
	GuiControl,2:,Brand,|
	COM_Init() ; 初始化COM
	prs := COM_CreateObject("ADODB.Recordset") ; 创建数据集 prs
	sSource  := "SELECT 品牌 from 品牌 where 产品类型='" . Class . "'" ; SQL查询语句
	com_invoke(prs, "open", sSource, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" . database) ; 打开数据库
	Loop { ; 查询
		if com_invoke(prs, "EOF") ; 当碰到EOF标志时，退出循环
			Break
		Brand := com_invoke(prs, "Fields[0].Value")    ; 获取数据集中第1个字段的值
		GuiControl,2:,Brand,%Brand%
		com_invoke(prs, "MoveNext")  ; 移到下一条
	}
	com_invoke(prs, "close") ; 关闭数据集
	com_release(prs)         ; 释放数据集
	COM_Term()
	Return
}
GetColor(database)
{
	LV_Delete()
	COM_Init() ; 初始化COM
	prs := COM_CreateObject("ADODB.Recordset") ; 创建数据集 prs
	sSource  := "SELECT 颜色 from 颜色" ; SQL查询语句
	com_invoke(prs, "open", sSource, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" . database) ; 打开数据库
	Loop { ; 查询
		if com_invoke(prs, "EOF") ; 当碰到EOF标志时，退出循环
			Break
		Color := com_invoke(prs, "Fields[0].Value")    ; 获取数据集中第1个字段的值
		Gui, 2:Default
		LV_Add("",Color)
		com_invoke(prs, "MoveNext")  ; 移到下一条
	}
	com_invoke(prs, "close") ; 关闭数据集
	com_release(prs)         ; 释放数据集
	COM_Term()
	Return,0
}
AddBrand(database,brand,class)
{
	com_init() ; 初始化COM
	pcon := COM_CreateObject("ADODB.Connection") ; 创建连接
	com_invoke(pcon, "open", "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" . database)     ; 打开连接
	SQL:="INSERT INTO 品牌(品牌,产品类型) values('" . brand . "','" . Class . "')"
	com_invoke(pcon, "Execute", SQL)  ; 执行 SQL语句
	com_invoke(pcon, "close")     ; 关闭连接
	com_term()
	return
}
AddColor(database,color)
{
	com_init() ; 初始化COM
	prs := COM_CreateObject("ADODB.Recordset") ; 创建数据集 prs
	sSource  := "SELECT MAX(编码) from 颜色"
	com_invoke(prs, "open", sSource, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" . database) ; 打开数据库
	Loop { ; 查询
		if com_invoke(prs, "EOF") ; 当碰到EOF标志时，退出循环
			Break
		MaxCode := com_invoke(prs, "Fields[0].Value")    ; 获取数据集中第1个字段的值
		com_invoke(prs, "MoveNext")  ; 移到下一条
	}
	com_invoke(prs, "close") ; 关闭数据集
	com_release(prs)         ; 释放数据集
	SetFormat,Float,03.0
	Code:=SubStr(MaxCode,2)+1.0
	Code:="C" Code
	pcon := COM_CreateObject("ADODB.Connection") ; 创建连接
	com_invoke(pcon, "open", "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" . database)     ; 打开连接
	SQL:="INSERT INTO 颜色(颜色,编码) values('" . Color . "','" . Code . "')"
	com_invoke(pcon, "Execute", SQL)  ; 执行 SQL语句
	com_invoke(pcon, "close")     ; 关闭连接
	com_term()
	return
}
GetColorCode(database,Color)
{
	COM_Init() ; 初始化COM
	prs := COM_CreateObject("ADODB.Recordset") ; 创建数据集 prs
	sSource  := "SELECT 编码 from 颜色 where 颜色='" . Color . "'" ; SQL查询语句
	com_invoke(prs, "open", sSource, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" . database) ; 打开数据库
	Loop { ; 查询
		if com_invoke(prs, "EOF") ; 当碰到EOF标志时，退出循环
			Break
		Code := com_invoke(prs, "Fields[0].Value")    ; 获取数据集中第1个字段的值
		com_invoke(prs, "MoveNext")  ; 移到下一条
	}
	com_invoke(prs, "close") ; 关闭数据集
	com_release(prs)         ; 释放数据集
	COM_Term()
	Return,Code
}
GetClassCode(database,class)
{
	COM_Init() ; 初始化COM
	prs := COM_CreateObject("ADODB.Recordset") ; 创建数据集 prs
	sSource  := "SELECT 主编号 from 产品 where 类型='" . Class . "'" ; SQL查询语句
	com_invoke(prs, "open", sSource, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" . database) ; 打开数据库
	Loop { ; 查询
		if com_invoke(prs, "EOF") ; 当碰到EOF标志时，退出循环
			Break
		Code := com_invoke(prs, "Fields[0].Value")    ; 获取数据集中第1个字段的值
		com_invoke(prs, "MoveNext")  ; 移到下一条
	}
	com_invoke(prs, "close") ; 关闭数据集
	com_release(prs)         ; 释放数据集
	COM_Term()
	Return,Code
}
GetMaxID(Database,Class)
{
	com_init() ; 初始化COM
	prs := COM_CreateObject("ADODB.Recordset") ; 创建数据集 prs
	sSource  := "SELECT MAX(编号) from " . Class
	com_invoke(prs, "open", sSource, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" . database) ; 打开数据库
	Loop { ; 查询
		if com_invoke(prs, "EOF") ; 当碰到EOF标志时，退出循环
			Break
		Max := com_invoke(prs, "Fields[0].Value")    ; 获取数据集中第1个字段的值
		com_invoke(prs, "MoveNext")  ; 移到下一条
	}
	com_invoke(prs, "close") ; 关闭数据集
	com_release(prs)         ; 释放数据集
	COM_Term()
	Return,Max
	}
CreatID(database,Class)
{
	GuiControlGet,Class,1:,Class
	if Class=
	{
		GuiControl,1:,Tip,请先选择一个商品分类
		Return
	}
	ClassCode:=GetClassCode(Database,Class)
	MaxID:=GetMaxID(Database,Class)
	StringReplace,MaxID,MaxID,%ClassCode%,,1
	MaxID++
	GuiControl,2:,ID,% ClassCode . MaxID
}
GetQueryItem(database,Class)
{
	if Class=
		Return
	GuiControl,1:,QueryItem,|SKU
	COM_Init() ; 初始化COM
	prs := COM_CreateObject("ADODB.Recordset") ; 创建数据集 prs
	sSource  := "SELECT * from " . Class ; SQL查询语句
	com_invoke(prs, "open", sSource, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" . database) ; 打开数据库

	COM_Error(0)
	Loop
	{ ; 查询
		q:="Fields[" . A_Index-1 "].Name"
		QueryItem := com_invoke(prs, q)    ; 获取数据集中第1个字段的值
		if QueryItem=颜色 || QueryItem=图片
			Continue
		if QueryItem=
			Break
		GuiControl,1:,QueryItem,%QueryItem%
	}
	com_invoke(prs, "close") ; 关闭数据集
	com_release(prs)         ; 释放数据集
	COM_Term()
}
Query(Database,SQL)
{
	COM_Init() ; 初始化COM
	prs := COM_CreateObject("ADODB.Recordset") ; 创建数据集 prs
	com_invoke(prs, "open", sql, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" . database) ; 打开数据库
	Loop { ; 查询
		if com_invoke(prs, "EOF") ; 当碰到EOF标志时，退出循环
			Break
		Item := com_invoke(prs, "Fields[0].Value")    ; 获取数据集中第1个字段的值
		Code:= com_invoke(prs, "Fields[1].Value")
		LV_Add("",Item,Code)
		com_invoke(prs, "MoveNext")  ; 移到下一条
	}
	com_invoke(prs, "close") ; 关闭数据集
	com_release(prs)         ; 释放数据集
	COM_Term()
	Return,0
}
QuerySKU(database,SKU)
{
	Sku:=RegExReplace(sku,"i)C\d{3}$","")
	if SKU=
		Return
	COM_Init() ; 初始化COM
	prs := COM_CreateObject("ADODB.Recordset") ; 创建数据集 prs
	sSource  := "SELECT 类型 from 产品" ; SQL查询语句
	com_invoke(prs, "open", sSource, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" . database) ; 打开数据库
	Loop { ; 查询
		if com_invoke(prs, "EOF") ; 当碰到EOF标志时，退出循环
			Break
		Class := com_invoke(prs, "Fields[0].Value")    ; 获取数据集中第1个字段的值
		SQL:="Select 产品名称,编号 From " . Class . " Where 编号 like '%" . sku . "%'"
		Query(Database,SQL)
		com_invoke(prs, "MoveNext")  ; 移到下一条
	}
	com_invoke(prs, "close") ; 关闭数据集
	com_release(prs)         ; 释放数据集
	COM_Term()
	Return
}
GetAllInfo(database,ID)
{
	if ID=
		Return
	if !RegExMatch(ID,"\w.*?(?=\d)",ClassCode)
		Return,-1
	COM_Init() ; 初始化COM
	prs := COM_CreateObject("ADODB.Recordset") ; 创建数据集 prs
	sql:="Select 类型 from 产品 where 主编号='" . ClassCode . "'"
	com_invoke(prs, "open", sql, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" . database) ; 打开数据库
	Loop { ; 查询
		if com_invoke(prs, "EOF") ; 当碰到EOF标志时，退出循环
			Break
		Class := com_invoke(prs, "Fields[0].Value")    ; 获取数据集中第1个字段的值
		com_invoke(prs, "MoveNext")  ; 移到下一条
	}
	com_invoke(prs, "close") ; 关闭数据集
	com_release(prs)         ; 释放数据集
	prs := COM_CreateObject("ADODB.Recordset") ; 创建数据集 prs
	sql:="Select * from " . Class . " where 编号='" . ID . "'"
	com_invoke(prs, "open", sql, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" . database) ; 打开数据库
	rtn=
	Loop { ; 查询
		if com_invoke(prs, "EOF") ; 当碰到EOF标志时，退出循环
			Break
		Com_error(0)
		Loop
		{
			if com_invoke(prs, "EOF") ; 当碰到EOF标志时，退出循环
				Break
			Key:=com_invoke(prs,"Fields[" . A_Index-1 . "].Name")
			Value := com_invoke(prs, "Fields[" . A_Index-1 . "].Value")    ; 获取数据集中第1个字段的值
			if key=
				break
			if key=颜色
				Value:=Code2Color(database,Value)
			rtn.=key . " : " . Value . "`n"
		}
		Com_error(1)
		com_invoke(prs, "MoveNext")  ; 移到下一条
	}
	com_invoke(prs, "close") ; 关闭数据集
	com_release(prs)
	COM_Term()
	StringTrimRight,rtn,rtn,1
	Return, rtn
}
Code2Color(database,CodeList)
{
	StringSplit,Code_,CodeList,/
	COM_Init() ; 初始化COM
	colorlist=
	Loop,% Code_0
	{
		prs := COM_CreateObject("ADODB.Recordset") ; 创建数据集 prs
		sql:="Select 颜色 from 颜色 where 编码='" . Code_%A_index% . "'"
		com_invoke(prs, "open", sql, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" . database) ; 打开数据库
		Color := com_invoke(prs, "Fields[0].Value")    ; 获取数据集中第1个字段的值
		ColorList.=Color . "/"
		com_invoke(prs, "close") ; 关闭数据集
		com_release(prs)
	}
	StringTrimRight,colorlist,colorlist,1
	COM_Term()
	Return,ColorList
}
Delete_ID(database,ID)
{
	ID:=RegExReplace(ID,"i)C\d{3}$","")
	if ID=
		Return,-1
	pos:=RegExMatch(ID,"[[:alpha:]]*?(?=\d)",ClassCode)
	if !pos
		Return,-2
	com_init() ; 初始化COM
	COM_ERROR(0)
	prs := COM_CreateObject("ADODB.Recordset") ; 创建数据集 prs
	sSource  := "SELECT 类型 from 产品 where 主编号='" . ClassCode . "'" ; SQL查询语句
	com_invoke(prs, "open", sSource, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" . database) ; 打开数据库
	Loop { ; 查询
		if com_invoke(prs, "EOF") ; 当碰到EOF标志时，退出循环
			Break
		Class := com_invoke(prs, "Fields[0].Value")    ; 获取数据集中第1个字段的值
		com_invoke(prs, "MoveNext")  ; 移到下一条
	}
	com_invoke(prs, "close") ; 关闭数据集
	com_release(prs)         ; 释放数据集

	pcon := COM_CreateObject("ADODB.Connection") ; 创建连接
	com_invoke(pcon, "open", "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" . database)     ; 打开连接
	SQL:="Delete from " . Class . " where 编号='" . ID . "'"
	com_invoke(pcon, "Execute", SQL)  ; 执行 SQL语句
	com_invoke(pcon, "close")     ; 关闭连接
	com_term()
	return
}
Delete_Color(database,Color)
{
	if Color=
		Return,-1
	com_init() ; 初始化COM
	COM_ERROR(0)
	pcon := COM_CreateObject("ADODB.Connection") ; 创建连接
	com_invoke(pcon, "open", "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" . database)     ; 打开连接
	SQL:="Delete from 颜色 where 颜色='" . Color . "'"
	com_invoke(pcon, "Execute", SQL)  ; 执行 SQL语句
	com_invoke(pcon, "close")     ; 关闭连接
	com_term()
	return
}
Delete_Brand(database,Brand)
{
	if Brand=
		Return,-1
	com_init() ; 初始化COM
	COM_ERROR(0)
	pcon := COM_CreateObject("ADODB.Connection") ; 创建连接
	com_invoke(pcon, "open", "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" . database)     ; 打开连接
	SQL:="Delete from 品牌 where 品牌='" . Brand . "'"
	com_invoke(pcon, "Execute", SQL)  ; 执行 SQL语句
	com_invoke(pcon, "close")     ; 关闭连接
	com_term()
	return
	}