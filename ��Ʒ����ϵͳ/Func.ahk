GetClass(database)
{
	COM_Init() ; ��ʼ��COM
	prs := COM_CreateObject("ADODB.Recordset") ; �������ݼ� prs
	sSource  := "SELECT ���� from ��Ʒ" ; SQL��ѯ���
	com_invoke(prs, "open", sSource, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" . database) ; �����ݿ�
	Loop { ; ��ѯ
		if com_invoke(prs, "EOF") ; ������EOF��־ʱ���˳�ѭ��
			Break
		Class := com_invoke(prs, "Fields[0].Value")    ; ��ȡ���ݼ��е�1���ֶε�ֵ
		GuiControl,1:,Class,%Class%
		com_invoke(prs, "MoveNext")  ; �Ƶ���һ��
	}
	com_invoke(prs, "close") ; �ر����ݼ�
	com_release(prs)         ; �ͷ����ݼ�
	COM_Term()
	Return
}
GetBrand(database,Class)
{
	if Class=
		Return
	GuiControl,2:,Brand,|
	COM_Init() ; ��ʼ��COM
	prs := COM_CreateObject("ADODB.Recordset") ; �������ݼ� prs
	sSource  := "SELECT Ʒ�� from Ʒ�� where ��Ʒ����='" . Class . "'" ; SQL��ѯ���
	com_invoke(prs, "open", sSource, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" . database) ; �����ݿ�
	Loop { ; ��ѯ
		if com_invoke(prs, "EOF") ; ������EOF��־ʱ���˳�ѭ��
			Break
		Brand := com_invoke(prs, "Fields[0].Value")    ; ��ȡ���ݼ��е�1���ֶε�ֵ
		GuiControl,2:,Brand,%Brand%
		com_invoke(prs, "MoveNext")  ; �Ƶ���һ��
	}
	com_invoke(prs, "close") ; �ر����ݼ�
	com_release(prs)         ; �ͷ����ݼ�
	COM_Term()
	Return
}
GetColor(database)
{
	LV_Delete()
	COM_Init() ; ��ʼ��COM
	prs := COM_CreateObject("ADODB.Recordset") ; �������ݼ� prs
	sSource  := "SELECT ��ɫ from ��ɫ" ; SQL��ѯ���
	com_invoke(prs, "open", sSource, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" . database) ; �����ݿ�
	Loop { ; ��ѯ
		if com_invoke(prs, "EOF") ; ������EOF��־ʱ���˳�ѭ��
			Break
		Color := com_invoke(prs, "Fields[0].Value")    ; ��ȡ���ݼ��е�1���ֶε�ֵ
		Gui, 2:Default
		LV_Add("",Color)
		com_invoke(prs, "MoveNext")  ; �Ƶ���һ��
	}
	com_invoke(prs, "close") ; �ر����ݼ�
	com_release(prs)         ; �ͷ����ݼ�
	COM_Term()
	Return,0
}
AddBrand(database,brand,class)
{
	com_init() ; ��ʼ��COM
	pcon := COM_CreateObject("ADODB.Connection") ; ��������
	com_invoke(pcon, "open", "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" . database)     ; ������
	SQL:="INSERT INTO Ʒ��(Ʒ��,��Ʒ����) values('" . brand . "','" . Class . "')"
	com_invoke(pcon, "Execute", SQL)  ; ִ�� SQL���
	com_invoke(pcon, "close")     ; �ر�����
	com_term()
	return
}
AddColor(database,color)
{
	com_init() ; ��ʼ��COM
	prs := COM_CreateObject("ADODB.Recordset") ; �������ݼ� prs
	sSource  := "SELECT MAX(����) from ��ɫ"
	com_invoke(prs, "open", sSource, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" . database) ; �����ݿ�
	Loop { ; ��ѯ
		if com_invoke(prs, "EOF") ; ������EOF��־ʱ���˳�ѭ��
			Break
		MaxCode := com_invoke(prs, "Fields[0].Value")    ; ��ȡ���ݼ��е�1���ֶε�ֵ
		com_invoke(prs, "MoveNext")  ; �Ƶ���һ��
	}
	com_invoke(prs, "close") ; �ر����ݼ�
	com_release(prs)         ; �ͷ����ݼ�
	SetFormat,Float,03.0
	Code:=SubStr(MaxCode,2)+1.0
	Code:="C" Code
	pcon := COM_CreateObject("ADODB.Connection") ; ��������
	com_invoke(pcon, "open", "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" . database)     ; ������
	SQL:="INSERT INTO ��ɫ(��ɫ,����) values('" . Color . "','" . Code . "')"
	com_invoke(pcon, "Execute", SQL)  ; ִ�� SQL���
	com_invoke(pcon, "close")     ; �ر�����
	com_term()
	return
}
GetColorCode(database,Color)
{
	COM_Init() ; ��ʼ��COM
	prs := COM_CreateObject("ADODB.Recordset") ; �������ݼ� prs
	sSource  := "SELECT ���� from ��ɫ where ��ɫ='" . Color . "'" ; SQL��ѯ���
	com_invoke(prs, "open", sSource, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" . database) ; �����ݿ�
	Loop { ; ��ѯ
		if com_invoke(prs, "EOF") ; ������EOF��־ʱ���˳�ѭ��
			Break
		Code := com_invoke(prs, "Fields[0].Value")    ; ��ȡ���ݼ��е�1���ֶε�ֵ
		com_invoke(prs, "MoveNext")  ; �Ƶ���һ��
	}
	com_invoke(prs, "close") ; �ر����ݼ�
	com_release(prs)         ; �ͷ����ݼ�
	COM_Term()
	Return,Code
}
GetClassCode(database,class)
{
	COM_Init() ; ��ʼ��COM
	prs := COM_CreateObject("ADODB.Recordset") ; �������ݼ� prs
	sSource  := "SELECT ����� from ��Ʒ where ����='" . Class . "'" ; SQL��ѯ���
	com_invoke(prs, "open", sSource, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" . database) ; �����ݿ�
	Loop { ; ��ѯ
		if com_invoke(prs, "EOF") ; ������EOF��־ʱ���˳�ѭ��
			Break
		Code := com_invoke(prs, "Fields[0].Value")    ; ��ȡ���ݼ��е�1���ֶε�ֵ
		com_invoke(prs, "MoveNext")  ; �Ƶ���һ��
	}
	com_invoke(prs, "close") ; �ر����ݼ�
	com_release(prs)         ; �ͷ����ݼ�
	COM_Term()
	Return,Code
}
GetMaxID(Database,Class)
{
	com_init() ; ��ʼ��COM
	prs := COM_CreateObject("ADODB.Recordset") ; �������ݼ� prs
	sSource  := "SELECT MAX(���) from " . Class
	com_invoke(prs, "open", sSource, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" . database) ; �����ݿ�
	Loop { ; ��ѯ
		if com_invoke(prs, "EOF") ; ������EOF��־ʱ���˳�ѭ��
			Break
		Max := com_invoke(prs, "Fields[0].Value")    ; ��ȡ���ݼ��е�1���ֶε�ֵ
		com_invoke(prs, "MoveNext")  ; �Ƶ���һ��
	}
	com_invoke(prs, "close") ; �ر����ݼ�
	com_release(prs)         ; �ͷ����ݼ�
	COM_Term()
	Return,Max
	}
CreatID(database,Class)
{
	GuiControlGet,Class,1:,Class
	if Class=
	{
		GuiControl,1:,Tip,����ѡ��һ����Ʒ����
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
	COM_Init() ; ��ʼ��COM
	prs := COM_CreateObject("ADODB.Recordset") ; �������ݼ� prs
	sSource  := "SELECT * from " . Class ; SQL��ѯ���
	com_invoke(prs, "open", sSource, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" . database) ; �����ݿ�

	COM_Error(0)
	Loop
	{ ; ��ѯ
		q:="Fields[" . A_Index-1 "].Name"
		QueryItem := com_invoke(prs, q)    ; ��ȡ���ݼ��е�1���ֶε�ֵ
		if QueryItem=��ɫ || QueryItem=ͼƬ
			Continue
		if QueryItem=
			Break
		GuiControl,1:,QueryItem,%QueryItem%
	}
	com_invoke(prs, "close") ; �ر����ݼ�
	com_release(prs)         ; �ͷ����ݼ�
	COM_Term()
}
Query(Database,SQL)
{
	COM_Init() ; ��ʼ��COM
	prs := COM_CreateObject("ADODB.Recordset") ; �������ݼ� prs
	com_invoke(prs, "open", sql, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" . database) ; �����ݿ�
	Loop { ; ��ѯ
		if com_invoke(prs, "EOF") ; ������EOF��־ʱ���˳�ѭ��
			Break
		Item := com_invoke(prs, "Fields[0].Value")    ; ��ȡ���ݼ��е�1���ֶε�ֵ
		Code:= com_invoke(prs, "Fields[1].Value")
		LV_Add("",Item,Code)
		com_invoke(prs, "MoveNext")  ; �Ƶ���һ��
	}
	com_invoke(prs, "close") ; �ر����ݼ�
	com_release(prs)         ; �ͷ����ݼ�
	COM_Term()
	Return,0
}
QuerySKU(database,SKU)
{
	Sku:=RegExReplace(sku,"i)C\d{3}$","")
	if SKU=
		Return
	COM_Init() ; ��ʼ��COM
	prs := COM_CreateObject("ADODB.Recordset") ; �������ݼ� prs
	sSource  := "SELECT ���� from ��Ʒ" ; SQL��ѯ���
	com_invoke(prs, "open", sSource, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" . database) ; �����ݿ�
	Loop { ; ��ѯ
		if com_invoke(prs, "EOF") ; ������EOF��־ʱ���˳�ѭ��
			Break
		Class := com_invoke(prs, "Fields[0].Value")    ; ��ȡ���ݼ��е�1���ֶε�ֵ
		SQL:="Select ��Ʒ����,��� From " . Class . " Where ��� like '%" . sku . "%'"
		Query(Database,SQL)
		com_invoke(prs, "MoveNext")  ; �Ƶ���һ��
	}
	com_invoke(prs, "close") ; �ر����ݼ�
	com_release(prs)         ; �ͷ����ݼ�
	COM_Term()
	Return
}
GetAllInfo(database,ID)
{
	if ID=
		Return
	if !RegExMatch(ID,"\w.*?(?=\d)",ClassCode)
		Return,-1
	COM_Init() ; ��ʼ��COM
	prs := COM_CreateObject("ADODB.Recordset") ; �������ݼ� prs
	sql:="Select ���� from ��Ʒ where �����='" . ClassCode . "'"
	com_invoke(prs, "open", sql, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" . database) ; �����ݿ�
	Loop { ; ��ѯ
		if com_invoke(prs, "EOF") ; ������EOF��־ʱ���˳�ѭ��
			Break
		Class := com_invoke(prs, "Fields[0].Value")    ; ��ȡ���ݼ��е�1���ֶε�ֵ
		com_invoke(prs, "MoveNext")  ; �Ƶ���һ��
	}
	com_invoke(prs, "close") ; �ر����ݼ�
	com_release(prs)         ; �ͷ����ݼ�
	prs := COM_CreateObject("ADODB.Recordset") ; �������ݼ� prs
	sql:="Select * from " . Class . " where ���='" . ID . "'"
	com_invoke(prs, "open", sql, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" . database) ; �����ݿ�
	rtn=
	Loop { ; ��ѯ
		if com_invoke(prs, "EOF") ; ������EOF��־ʱ���˳�ѭ��
			Break
		Com_error(0)
		Loop
		{
			if com_invoke(prs, "EOF") ; ������EOF��־ʱ���˳�ѭ��
				Break
			Key:=com_invoke(prs,"Fields[" . A_Index-1 . "].Name")
			Value := com_invoke(prs, "Fields[" . A_Index-1 . "].Value")    ; ��ȡ���ݼ��е�1���ֶε�ֵ
			if key=
				break
			if key=��ɫ
				Value:=Code2Color(database,Value)
			rtn.=key . " : " . Value . "`n"
		}
		Com_error(1)
		com_invoke(prs, "MoveNext")  ; �Ƶ���һ��
	}
	com_invoke(prs, "close") ; �ر����ݼ�
	com_release(prs)
	COM_Term()
	StringTrimRight,rtn,rtn,1
	Return, rtn
}
Code2Color(database,CodeList)
{
	StringSplit,Code_,CodeList,/
	COM_Init() ; ��ʼ��COM
	colorlist=
	Loop,% Code_0
	{
		prs := COM_CreateObject("ADODB.Recordset") ; �������ݼ� prs
		sql:="Select ��ɫ from ��ɫ where ����='" . Code_%A_index% . "'"
		com_invoke(prs, "open", sql, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" . database) ; �����ݿ�
		Color := com_invoke(prs, "Fields[0].Value")    ; ��ȡ���ݼ��е�1���ֶε�ֵ
		ColorList.=Color . "/"
		com_invoke(prs, "close") ; �ر����ݼ�
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
	com_init() ; ��ʼ��COM
	COM_ERROR(0)
	prs := COM_CreateObject("ADODB.Recordset") ; �������ݼ� prs
	sSource  := "SELECT ���� from ��Ʒ where �����='" . ClassCode . "'" ; SQL��ѯ���
	com_invoke(prs, "open", sSource, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" . database) ; �����ݿ�
	Loop { ; ��ѯ
		if com_invoke(prs, "EOF") ; ������EOF��־ʱ���˳�ѭ��
			Break
		Class := com_invoke(prs, "Fields[0].Value")    ; ��ȡ���ݼ��е�1���ֶε�ֵ
		com_invoke(prs, "MoveNext")  ; �Ƶ���һ��
	}
	com_invoke(prs, "close") ; �ر����ݼ�
	com_release(prs)         ; �ͷ����ݼ�

	pcon := COM_CreateObject("ADODB.Connection") ; ��������
	com_invoke(pcon, "open", "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" . database)     ; ������
	SQL:="Delete from " . Class . " where ���='" . ID . "'"
	com_invoke(pcon, "Execute", SQL)  ; ִ�� SQL���
	com_invoke(pcon, "close")     ; �ر�����
	com_term()
	return
}
Delete_Color(database,Color)
{
	if Color=
		Return,-1
	com_init() ; ��ʼ��COM
	COM_ERROR(0)
	pcon := COM_CreateObject("ADODB.Connection") ; ��������
	com_invoke(pcon, "open", "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" . database)     ; ������
	SQL:="Delete from ��ɫ where ��ɫ='" . Color . "'"
	com_invoke(pcon, "Execute", SQL)  ; ִ�� SQL���
	com_invoke(pcon, "close")     ; �ر�����
	com_term()
	return
}
Delete_Brand(database,Brand)
{
	if Brand=
		Return,-1
	com_init() ; ��ʼ��COM
	COM_ERROR(0)
	pcon := COM_CreateObject("ADODB.Connection") ; ��������
	com_invoke(pcon, "open", "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" . database)     ; ������
	SQL:="Delete from Ʒ�� where Ʒ��='" . Brand . "'"
	com_invoke(pcon, "Execute", SQL)  ; ִ�� SQL���
	com_invoke(pcon, "close")     ; �ر�����
	com_term()
	return
	}