<%
Session("DebugInfo")=""

Sub debug(info)
	Session("DebugInfo")=Session("DebugInfo")&vbNewLine&info
End Sub
Sub LoadingPage()
	'**************************************************************
	'Author:华工ASP开发小组
	'Name:打开ASP页面时 显示 装载页面信息
	'Created Date:	2004/11/04
	'Purpose:服务器运行ASP文件时 强制把 装载信息发到客户端显示 
	'Relate pages:与 EndPage()并用 该函数一般放在ASP文件头
	'Modify By:	Dajing
	'Modify date:	2004/11/04
	'**************************************************************
	Response.Write("<div id='LoadingPage' >")&vbcrlf
	Response.Write("_")&vbcrlf
	Response.Write("</div>")&vbcrlf
	Response.Write("<script>LoadingPage.innerText = '';</script>")&vbcrlf
	Response.Write("<script language=javascript>;")&vbcrlf
	Response.Write("var dots = 0;var dotmax = 10;function ShowWait()")&vbcrlf
	Response.Write("{var output; output = '正在装载页面';dots++;if(dots>=dotmax)dots=1;")&vbcrlf
	Response.Write("for(var x = 0;x < dots;x++){output += '·';}LoadingPage.innerText =  output;}")&vbcrlf
	Response.Write("function StartShowWait(){LoadingPage.style.visibility = 'visible'; ")&vbcrlf
	Response.Write("window.setInterval('ShowWait()',1000);}")&vbcrlf
	Response.Write("function HideWait(){LoadingPage.style.display='none';")&vbcrlf
	Response.Write("window.clearInterval();}")&vbcrlf
	Response.Write("StartShowWait();</script>")&vbcrlf
	Response.Flush		'强制把 装载信息发到客户端显示
End Sub
Sub EndPage()
	'**************************************************************
	'Author:华工ASP开发小组
	'Name:关闭 函数LoadingPage() 显示出的提示信息
	'Created Date:	2004/11/04
	'Purpose:与函数并用LoadingPage()
	'Relate pages:一般该函数放在ASP文件末端
	'Modify By:	Dajing
	'Modify date:	2004/11/04
	'**************************************************************
	Response.write "<SCRIPT LANGUAGE=""JavaScript"">"&vbcrlf
	Response.write "HideWait()"&vbcrlf				'该函数在LoadingPage 用来关闭显示出来的信息
	'Response.write "LoadingPage.style.display=""none"""&vbcrlf
	Response.write "</SCRIPT>"&vbcrlf
End Sub

Sub Save_Visit_UserIP(Visit_Module,Visit_UserIp,Visit_UserNo,Visit_FrontURL)					'by Dajing
'该过程 是记录访问者IP 到表Visit_Info
'Visit_Ip  访问者IP
'Visit_UserNo 访问者 登陆名(可从SESSION处获得)
'Visit_Module 访问者 访问的模块
	Dim InsStr,conn

	Connect conn

	If Visit_UserIp="" or Visit_UserNo="" or Visit_Module="" or Visit_FrontURL="" Then
		Response.write "<SCRIPT LANGUAGE=""JavaScript"">"
		Response.write "alert('超时，请重新登陆');"
		Response.write "window.location.href='../admin/err/timeout.asp'"
		Response.write "</SCRIPT>"
		Response.end
	Else
		InsStr="Insert Into Visit_Info (Visit_UserIp,Visit_UserNo,Visit_Module,Visit_FrontURL) VALUES ("&"'"&Visit_UserIp&"','"&Visit_UserNo&"','"&Visit_Module&"','"&Visit_FrontURL&"')"
	End If 
	'Response.write InsStr
	'Response.end
	conn.Execute InsStr
End Sub

'=====================================================
Sub Get_ListJavaMenu(ArrayList,j,FormName,Pub_Term)					'by Dajing
		'该过程的功能是 关联下拉菜单（JAVASCIPT方式的）
		'Pub_Term 为公共条件，每次刷新都会用到
		'FormName 为下拉框所在FORM 的名字
		'ArrayList 是传入的一数组  j为传入数组下标的值
		'例如 ArrayList数组有下例值 每个单元分别表示为
		'ArrayList(0,0)="类型"					'显示的名字
		'ArrayList(0,1)=Table					'表名
		'ArrayList(0,2)="TEACHTYPE_ID"			'SELECT VALUE的值
		'ArrayList(0,3)="TEACHTYPE_NAME"			'SELECT 显示出来的值
		'ArrayList(0,4)="4"						'SELECT 默认值
		'ArrayList(0,5)=""						'SELECT 默认的条件值

		'第一次运行该过程时HTML 要运行 JAVASCRIPT 的过程 
			'例：<BODY onload="return On_Load()"> 

		'选定 后运行 JAVASCRIPT函数Chk_Select() 后 SELECT的最终值就可存放在该SELEC变量名加 'In_'的TEXT变量上 
		'例: SELECT的NAME为TEACHTYPE_ID 那么选定后该SELECT的最终VALUE会放在In_TEACHTYPE_ID里
Dim SelectLen,SelectLenStr,SelectLenStrOld,RsSqlStr,FieldId,Field
Dim bNoCheck
Dim conn

bNoCheck=Right(FormName,8)="_nocheck"
SelectLen=10
SelectLenStr=""
FieldId=""
Field=""
Connect con

'-----------------把SELECT 里的东东从数据库里倒到JAVASCRIPT的数组里--------------
Response.write "<script language=""JavaScript"">"&vbcrlf
For i=0 to j

	SelectLenNew=SelectLen*(i+1)

	For ii=1 To SelectLen
		SelectLenStr=SelectLenStr&"0"
	Next
	FieldId="replace(right('"&SelectLenStr&"',"&SelectLenNew&"-len(convert(char("&SelectLenNew&"),"&ArrayList(i,2)&"))) + convert(char("&SelectLenNew&"),"&ArrayList(i,2)&"),' ','')"
	If i<>0 Then
		Field=Field&" + "
	End If
	Field=Field&FieldId

	RsSqlStr="Select DISTINCT "&ArrayList(i,2)&"="&Field&","&ArrayList(i,3)&" From "&ArrayList(i,1)&" where 1=1"&ArrayList(i,5)&Pub_Term&" order by "&ArrayList(i,3)
	If InStr(UCase(ArrayList(i,3)),"YEAR")>0 Or UCase(ArrayList(i,3))="PERIOD_NAME" Then RsSqlStr=RsSqlStr&" desc"
	GetRecordSetNoLock conn,rsSelMenu,RsSqlStr,resultSelMenu

'----------------把数据库里边的东东写到JAVA脚本，把数据内容传JAVA数组里------------
	Response.write vbcrlf&"var "&ArrayList(i,2)&"Code = new Array("&resultSelMenu&");"&vbcrlf
	Response.write "var "&ArrayList(i,2)&"Desc = new Array("&resultSelMenu&");"&vbcrlf

	Response.write ArrayList(i,2)&"Code[0] = '"&SelectLenStr&"';"&vbcrlf
	Response.write ArrayList(i,2)&"Desc[0] = '请选择"&ArrayList(i,0)&"';"&vbcrlf

	For ii=1 To resultSelMenu
		Response.write ArrayList(i,2)&"Code["&ii&"] = '"&rsSelMenu(0)&"';"&vbcrlf
		Response.write ArrayList(i,2)&"Desc["&ii&"] = '"&GetPartString(rsSelMenu(1),15)&"';"&vbcrlf
		rsSelMenu.MoveNext
	Next
Next

For i=0 To j-1
	'调用控制下拉框代码的FUNCTION
	Response.write "function Sel"&ArrayList(i,2)&"_OnChange(){"&vbcrlf
	Response.write "var fm=document.getElementById('"&FormName&"');"&vbcrlf
	
	'SelectLenNew 是下拉框变量的长度 是提供下一函数substr(0,intLength)中的intLength使用
	'得到SelectLenNew的公式是： 
		'----------------------
	SelectLenNew=0
	For ii=0 to i
		SelectLenNew=SelectLenNew+SelectLen*(ii)	'计算上一次 SelectLenNew的值
	Next
'	ii=i+1
	For ii=i+1 To j
		SelectLenNew=SelectLenNew+SelectLen*(ii)	'计算这一次 SelectLenNew的值
		'----------------------
		Response.write "Change"&ArrayList(ii,2)&"(fm."&ArrayList(ii-1,2)&".value,"&SelectLenNew&");"&vbcrlf
	Next
	Response.write "}//end"&i&vbcrlf&vbcrlf
Next

SelectLenStr=""
For i=1 to j
	'下边的JAVA函数 是控制下拉框的代码(起来联动)
	Response.write "function Change"&ArrayList(i,2)&"(strValue,intLength){"&vbcrlf
	Response.write "var fm=document.getElementById('"&FormName&"');"&vbcrlf
	Response.write "fm."&ArrayList(i,2)&".length = 0;"&vbcrlf
	Response.write "var strCode;"&vbcrlf
	Response.write "var k = 0;"&vbcrlf
	Response.write "var i = 0;"&vbcrlf
	Response.write "for (i = 0;i<"&ArrayList(i,2)&"Code.length;i++){"&vbcrlf
	Response.write "strCode = "&ArrayList(i,2)&"Code[i].substr(0,intLength);"&vbcrlf
	Response.write "if ((strValue==strCode) || (i==0)){"&vbcrlf
	Response.write "fm."&ArrayList(i,2)&".options[k] = new Option("&ArrayList(i,2)&"Desc[i],"&ArrayList(i,2)&"Code[i]);"&vbcrlf

	If i<>j Then
		If ArrayList(i,4)<>"" Then
			Response.write "	if("&ArrayList(i,2)&"Code[i].substr(intLength).ltrim()"&"=='"&ArrayList(i,4)&"'){"&vbcrlf	'判断当前值是否为默认值
			Response.write "		fm."&ArrayList(i,2)&".selectedIndex = k;"&vbcrlf	'设置默认值
	  	Response.write "		var onchange=Sel"&ArrayList(i,2)&"_OnChange; if(onchange) onchange();"&vbcrlf	'更新下一个下拉框的内容
			Response.write "	}//if"&i&vbcrlf
		Else
			Response.write "	if(k==1){"&vbcrlf
			Response.write "		fm."&ArrayList(i,2)&".selectedIndex = 1;"&vbcrlf	'设置默认值
	  	Response.write "		var onchange=Sel"&ArrayList(i,2)&"_OnChange; if(onchange) onchange();"&vbcrlf	'更新下一个下拉框的内容
			Response.write "	}"&vbcrlf
		End If
	End If

	Response.write "k = k + 1;"&vbcrlf
	Response.write "}//if"&vbcrlf
	Response.write "}//for"&vbcrlf

	Response.write "}//end"&i&vbcrlf&vbcrlf
Next

'第一次运行该过程时 把SELECT初初化的JAVA函数
'例<BODY onload="return On_Load()">
	Response.write "function On_Load(){"&vbcrlf
	Response.write "var fm=document.getElementById('"&FormName&"');"&vbcrlf
	Response.write "var i;"&vbcrlf
	i=0
	Response.write "for (i = 0;i<"&ArrayList(i,2)&"Code.length;i++){"&vbcrlf
	Response.write "fm."&ArrayList(i,2)&".options[i] = new Option("&ArrayList(i,2)&"Desc[i],"&ArrayList(i,2)&"Code[i]);"&vbcrlf
	If ArrayList(i,4)<>"" Then
		Response.write "	if("&ArrayList(i,2)&"Code[i].ltrim()"&"=='"&ArrayList(i,4)&"'){"&vbcrlf '判断当前值是否为默认值
		Response.write "		fm."&ArrayList(i,2)&".selectedIndex = i;"&vbcrlf	'设置默认值
	  Response.write "		var onchange=Sel"&ArrayList(i,2)&"_OnChange; if(onchange) onchange();"&vbcrlf	'更新下一个下拉框的内容
		Response.write "	}//if"&vbcrlf
	Else
		Response.write "	if(i==1){"&vbcrlf
		Response.write "	fm."&ArrayList(i,2)&".selectedIndex = 1;"&vbcrlf	'设置默认值
	  Response.write "		var onchange=Sel"&ArrayList(i,2)&"_OnChange; if(onchange) onchange();"&vbcrlf	'更新下一个下拉框的内容
		Response.write "	}"&vbcrlf
	End If

	Response.write "}//for"&vbcrlf
	Response.write "}//end"&vbcrlf&vbcrlf

	'下面的JAVA函数是判断 提交时SELECT是否选择
	Response.write "function Chk_Select(){"&vbcrlf
	Response.write "var fm=document.getElementById('"&FormName&"');"&vbcrlf
	StartLen=0
	EndLen=0
	For i=0 To j
		If Not bNoCheck Then
			'判断
			If i<>j and (i+1)<>j Then	'最后一个与倒数第二个SELECT不判断 如果要用户请自行加语句判断
				Response.write "if(fm."&ArrayList(i,2)&".selectedIndex=='0'){"&vbcrlf
				Response.write "alert('请选择"&ArrayList(i,0)&"');"&vbcrlf
				Response.write "fm."&ArrayList(i,2)&".focus();"&vbcrlf
				Response.write "return false;"&vbcrlf
				Response.write "}//if"&i&vbcrlf
			End If
		End If
		'取结果
		StartLen=StartLen+SelectLen*i
		EndLen=EndLen+SelectLen*(i+1)

		Response.write "fm.In_"&ArrayList(i,2)&".value=fm."&ArrayList(i,2)&".value.substr("&StartLen&","&EndLen&").ltrim();"&vbcrlf

	Next
	Response.write "return true;"&vbcrlf
	Response.write "}//end"&i&vbcrlf&vbcrlf

	Response.write "String.prototype.ltrim = function()    //取字符串左边零函数"&vbcrlf
	Response.write "{"&vbcrlf
	Response.write "    // 用正则表达式将字符串左边零"&vbcrlf
	Response.write "    // 用空字符串替代。"&vbcrlf
	Response.write "    return this.replace(/(^0*)/g, """");"&vbcrlf
	Response.write "}"&vbcrlf&vbcrlf

Response.write "</SCRIPT>"&vbcrlf&vbcrlf


'-----------------------调用数据库把SELECT 里边的选择列出来------------------------
SelectLenStr=""
FieldId=""
Field=""
For  i=0 To j
	If ((i+1) mod 4)=1 Then
		Response.write "<tr bgcolor=""ghostwhite"">"
	End If

	Response.write "<td>"&ArrayList(i,0)&"</td>"&vbcrlf
	Response.write "<td"
	Response.write ">"
	Response.write "<input type='hidden' name='In_"&ArrayList(i,2)&"'><!-- 该变量是存放结果（select）的值 -->"&vbcrlf
	Response.write "<select name='"&ArrayList(i,2)&"' "
	If i<>j Then
		'最后一个不用运行JAVASCRIPT
		Response.write "onChange=""javascript:Sel"&ArrayList(i,2)&"_OnChange();"""
	End If
	Response.write ">"&vbcrlf

	Response.write 	"</select>"
	Response.write "</td>"&vbcrlf
	If ((i+1) mod 3)=0 and i<>2 Then	'or i=j
		Response.write "</tr>"&vbcrlf
	End If
Next
' 最后的 TR没有放上</tr> 是因为有时会把按纽放上来 如果不放在调用函数后加上</tr>即可
CloseConn conn
CloseRs rsSelMenu
End Sub

Sub Get_ListJavaMenu_More(ArrayList,j,FormName,Pub_Term,Snum)					'by Dajing
		'该过程的功能是 关联下拉菜单（JAVASCIPT方式的）
		'Snum 给该过程生成的JAVASCRIPT加上一个标识~以便在同一页面多次调用该过程 By xiaopeng
		'Pub_Term 为公共条件，每次刷新都会用到
		'FormName 为下拉框所在FORM 的名字
		'ArrayList 是传入的一数组  j为传入数组下标的值
		'例如 ArrayList数组有下例值 每个单元分别表示为
		'ArrayList(0,0)="类型"					'显示的名字
		'ArrayList(0,1)=Table					'表名
		'ArrayList(0,2)="TEACHTYPE_ID"			'SELECT VALUE的值
		'ArrayList(0,3)="TEACHTYPE_NAME"			'SELECT 显示出来的值
		'ArrayList(0,4)="4"						'SELECT 默认值
		'ArrayList(0,5)=""						'SELECT 默认的条件值

		'第一次运行该过程时HTML 要运行 JAVASCRIPT 的过程 
			'例：<BODY onload="return On_Load()"> 

		'选定 后运行 JAVASCRIPT函数Chk_Select() 后 SELECT的最终值就可存放在该SELEC变量名加 'In_'的TEXT变量上 
		'例: SELECT的NAME为TEACHTYPE_ID 那么选定后该SELECT的最终VALUE会放在In_TEACHTYPE_ID里
Dim SelectLen,SelectLenStr,SelectLenStrOld,RsSqlStr,FieldId,Field
Dim conn

SelectLen=10
SelectLenStr=""
FieldId=""
Field=""
Connect conn

'-----------------把SELECT 里的东东从数据库里倒到JAVASCRIPT的数组里--------------
Response.write "<script language=""JavaScript"">"&vbcrlf
For i=0 to j

	SelectLenNew=SelectLen*(i+1)

	For ii=1 To SelectLen
		SelectLenStr=SelectLenStr&"0"
	Next
	FieldId="replace(right('"&SelectLenStr&"',"&SelectLenNew&"-len(convert(char("&SelectLenNew&"),"&ArrayList(i,2)&"))) + convert(char("&SelectLenNew&"),"&ArrayList(i,2)&"),' ','')"
	If i<>0 Then
		Field=Field&" + "
	End If
	Field=Field&FieldId

	RsSqlStr="Select DISTINCT "&ArrayList(i,2)&"="&Field&","&ArrayList(i,3)&" From "&ArrayList(i,1)&" where 1=1"&ArrayList(i,5)&Pub_Term
	GetRecordSetNoLock conn,rsSelMenu,RsSqlStr,resultSelMenu

'----------------把数据库里边的东东写到JAVA脚本，把数据内容传JAVA数组里------------
	Response.write vbcrlf&"var "&ArrayList(i,2)&Snum&"Code = new Array("&resultSelMenu&");"&vbcrlf
	Response.write "var "&ArrayList(i,2)&Snum&"Desc = new Array("&resultSelMenu&");"&vbcrlf

	Response.write ArrayList(i,2)&Snum&"Code[0] = '"&SelectLenStr&"';"&vbcrlf
	Response.write ArrayList(i,2)&Snum&"Desc[0] = '请选择"&ArrayList(i,0)&"';"&vbcrlf

	For ii=1 To resultSelMenu
		Response.write ArrayList(i,2)&Snum&"Code["&ii&"] = '"&rsSelMenu(0)&"';"&vbcrlf
		Response.write ArrayList(i,2)&Snum&"Desc["&ii&"] = '"&rsSelMenu(1)&"';"&vbcrlf
		rsSelMenu.MoveNext
	Next
Next

For i=0 To j				'如果需要最后一个不用默认值 这行加上 -1
	'调用控制下拉框代码的FUNCTION
	Response.write "function Sel"&ArrayList(i,2)&Snum&"_OnChange(){"&vbcrlf
	Response.write "var fm=document.getElementById('"&FormName&"');"&vbcrlf
	
	'SelectLenNew 是下拉框变量的长度 是提供下一函数substr(0,intLength)中的intLength使用
	'得到SelectLenNew的公式是： 
		'----------------------
	SelectLenNew=0
	For ii=0 to i
		SelectLenNew=SelectLenNew+SelectLen*(ii)	'计算上一次 SelectLenNew的值
	Next
'	ii=i+1
	For ii=i+1 To j
		SelectLenNew=SelectLenNew+SelectLen*(ii)	'计算这一次 SelectLenNew的值
		'----------------------
		Response.write "Change"&ArrayList(ii,2)&Snum&"(fm."&ArrayList(ii-1,2)&Snum&".value,"&SelectLenNew&");"&vbcrlf
	Next
	Response.write "}//end"&i&vbcrlf&vbcrlf
Next

SelectLenStr=""
For i=1 to j
	'下边的JAVA函数 是控制下拉框的代码(起来联动)
	Response.write "function Change"&ArrayList(i,2)&Snum&"(strValue,intLength){"&vbcrlf
	Response.write "var fm=document.getElementById('"&FormName&"');"&vbcrlf
	Response.write "fm."&ArrayList(i,2)&Snum&".length = 0;"&vbcrlf
	Response.write "var strCode;"&vbcrlf
	Response.write "var k = 0;"&vbcrlf
	Response.write "var i = 0;"&vbcrlf
	Response.write "for (i = 0;i<"&ArrayList(i,2)&Snum&"Code.length;i++){"&vbcrlf
	Response.write "strCode = "&ArrayList(i,2)&Snum&"Code[i].substr(0,intLength);"&vbcrlf
	Response.write "if ((strValue==strCode) || (i==0)){"&vbcrlf
	Response.write "fm."&ArrayList(i,2)&Snum&".options[k] = new Option("&ArrayList(i,2)&Snum&"Desc[i],"&ArrayList(i,2)&Snum&"Code[i]);"&vbcrlf

	'If i<>j Then	'如果需要最后一个不用默认值 这行就可注释掉
		If ArrayList(i,4)<>"" Then
			Response.write "	if("&ArrayList(i,2)&Snum&"Code[i].substr(intLength).ltrim()"&"=='"&ArrayList(i,4)&"'){"&vbcrlf	'判断当前值是否为默认值
			Response.write "		fm."&ArrayList(i,2)&Snum&".selectedIndex = k;"&vbcrlf	'设置默认值
			Response.write "		Sel"&ArrayList(i,2)&Snum&"_OnChange(); "&vbcrlf	'更新下一个下拉框的内容
			Response.write "	}//if"&i&vbcrlf
		Else
			Response.write "	if(k==1){"&vbcrlf
			Response.write "		fm."&ArrayList(i,2)&Snum&".selectedIndex = 1;"&vbcrlf	'设置默认值
			Response.write "		Sel"&ArrayList(i,2)&Snum&"_OnChange(); "&vbcrlf	'更新下一个下拉框的内容
			Response.write "	}"&vbcrlf
		End If
	'End If

	Response.write "k = k + 1;"&vbcrlf
	Response.write "}//if"&vbcrlf
	Response.write "}//for"&vbcrlf

	Response.write "}//end"&i&vbcrlf&vbcrlf
Next

'第一次运行该过程时 把SELECT初初化的JAVA函数
'例<BODY onload="return On_Load()">
	Response.write "function On_Load"&Snum&"(){"&vbcrlf
	Response.write "var fm=document.getElementById('"&FormName&"');"&vbcrlf
	Response.write "var i = 0;"&vbcrlf
	i=0
	Response.write "for (i = 0;i<"&ArrayList(i,2)&Snum&"Code.length;i++){"&vbcrlf
	Response.write "fm."&ArrayList(i,2)&Snum&".options[i] = new Option("&ArrayList(i,2)&Snum&"Desc[i],"&ArrayList(i,2)&Snum&"Code[i]);"&vbcrlf
	If ArrayList(i,4)<>"" Then
		Response.write "	if("&ArrayList(i,2)&Snum&"Code[i].ltrim()"&"=='"&ArrayList(i,4)&"'){"&vbcrlf '判断当前值是否为默认值
		Response.write "		fm."&ArrayList(i,2)&Snum&".selectedIndex = i;"&vbcrlf	'设置默认值
		Response.write "		Sel"&ArrayList(i,2)&Snum&"_OnChange(); "&vbcrlf	'更新下一个下拉框的内容
		Response.write "	}//if"&vbcrlf
	Else
		Response.write "	if(i==1){"&vbcrlf
		Response.write "	fm."&ArrayList(i,2)&Snum&".selectedIndex = 1;"&vbcrlf	'设置默认值
		Response.write "	Sel"&ArrayList(i,2)&Snum&"_OnChange(); "&vbcrlf	'更新下一个下拉框的内容
		Response.write "	}"&vbcrlf
	End If

	Response.write "}//for"&i&vbcrlf
	Response.write "}//end"&i&vbcrlf&vbcrlf


'下面的JAVA函数是判断 提交时SELECT是否选择
	Response.write "function Chk_Select"&Snum&"(){"&vbcrlf
	Response.write "var fm=document.getElementById('"&FormName&"');"&vbcrlf
	StartLen=0
	EndLen=0
	For i=0 To j
		'判断
		If i<>j and (i+1)<>j Then	'最后一个与倒数第二个SELECT不判断 如果要用户请自行加语句判断
			Response.write "if(fm."&ArrayList(i,2)&Snum&".selectedIndex=='0'){"&vbcrlf
			Response.write "alert('请选择"&ArrayList(i,0)&"');"&vbcrlf
			Response.write "fm."&ArrayList(i,2)&Snum&".focus();"&vbcrlf
			Response.write "return false;"&vbcrlf
			Response.write "}//if"&i&vbcrlf
		End If
		'取结果
		StartLen=StartLen+SelectLen*i
		EndLen=EndLen+SelectLen*(i+1)

		Response.write "fm.In_"&ArrayList(i,2)&Snum&".value=fm."&ArrayList(i,2)&Snum&".value.substr("&StartLen&","&EndLen&").ltrim();"&vbcrlf

	Next
	Response.write "return true;"&vbcrlf
	Response.write "}//end"&i&vbcrlf&vbcrlf

	Response.write "String.prototype.ltrim = function()    //取字符串左边零函数"&vbcrlf
	Response.write "{"&vbcrlf
	Response.write "    // 用正则表达式将字符串左边零"&vbcrlf
	Response.write "    // 用空字符串替代。"&vbcrlf
	Response.write "    return this.replace(/(^0*)/g, """");"&vbcrlf
	Response.write "}"&vbcrlf&vbcrlf

Response.write "</SCRIPT>"&vbcrlf&vbcrlf


'-----------------------调用数据库把SELECT 里边的选择列出来------------------------
SelectLenStr=""
FieldId=""
Field=""
For  i=0 To j
	If ((i+1) mod 3)=1 Then
		Response.write "<tr bgcolor=""ghostwhite"">"
	End If

	Response.write "<td width=50>"&ArrayList(i,0)&"</td>"&vbcrlf
	Response.write "<td"
	Response.write ">"
	Response.write "<input type='hidden' name='In_"&ArrayList(i,2)&Snum&"'><!-- 该变量是存放结果（select）的值 -->"&vbcrlf
	Response.write "<select name='"&ArrayList(i,2)&Snum&"' "
	If i<>j Then
		'最后一个不用运行JAVASCRIPT
		Response.write "onChange=""javascript:Sel"&ArrayList(i,2)&Snum&"_OnChange();"""
	End If
	Response.write ">"&vbcrlf

	Response.write 	"</select>"
	Response.write "</td>"&vbcrlf
	If ((i+1) mod 3)=0 and i<>2 Then	'or i=j
		Response.write "</tr>"&vbcrlf
	End If
Next
' 最后的 TR没有放上</tr> 是因为有时会把按纽放上来 如果不放在调用函数后加上</tr>即可
CloseConn conn
CloseRs rsSelMenu
End Sub

Function GetPartString(Str,SLen)
	If len(Str)>SLen Then
		GetPartString=mid(Str,1,SLen)&"..."
	Else
		GetPartString=Str
	End If
End Function

Sub back(ValueName,Value,msg,ReturnURL,ReturnType)
'**************************************************************
'Author:华工ASP开发小组
'Name:通用返回过程
'Created Date:	2005/1/11
'Purpose:通过传惨 ValueName,Value,ReturnURL生成HTML 后自动返回
'Relate pages:多用于保存、修改、删除后的页面调用
'Modify By:	Dajing
'Modify date:	2005/3/3

'ValueName 返回变更集字符串  
'	例：返回变量名为 Plan_Year与TeachType_Id时 该变量的值为"Plan_Year,TeachType_Id" 值与值之间用逗号分隔
'Value 返回变量值集字符串  
'	例：返回变量值为 2004至2005与3时 该变量的值为"2004至2005,3" 值与值之间用逗号分隔
'ReturnURL 返回URL地址
'ReturnType 调用返回过程类型 为html则为单独页面，值为其它时表示与其它FORM共存
'**************************************************************
	dim i
	ValueName=split(ValueName,",")
	Value=split(Value,",")
	If ReturnURL="" Then ReturnURL=Request.ServerVariables("HTTP_REFERER")

	Response.write vbcrlf
	If ucase(ReturnType)="HTML" Then
		'-----HTML文件头------
		Response.write "<HTML>"&vbcrlf
		Response.write "<HEAD>"&vbcrlf
		Response.write "<TITLE></TITLE>"&vbcrlf
		Response.write "</HEAD>"&vbcrlf
		Response.write "<BODY onLoad=""Retu_Load()"">"&vbcrlf
	End If

	'------FORM------
	Response.write "<form id=""ReturnForm"" METHOD=POST ACTION="""&ReturnURL&""">"&vbcrlf&vbcrlf
	For i=0 To ubound(ValueName)
		Response.write "	<input type=""hidden"" name="""&ValueName(i)&""" value="""&Value(i)&""">"&vbcrlf
	Next
	Response.write "</FORM>"&vbcrlf
	'------FORM------

	If ucase(ReturnType)="HTML" Then
		'-----HTML文件尾------
		Response.write "</BODY>"&vbcrlf
		Response.write "</HTML>"&vbcrlf&vbcrlf
	End If

	Response.write "<SCRIPT LANGUAGE=""JavaScript"">"&vbcrlf
	Response.write "	function Retu_Load()"&vbcrlf
	Response.write "	{"&vbcrlf
	If ucase(ReturnType)="HTML" Then
		Response.write "		alert('"&msg&"');"&vbcrlf
	End If
	Response.write "		document.all.ReturnForm.submit();"&vbcrlf
	Response.write "	}"&vbcrlf
	Response.write "</SCRIPT>"&vbcrlf
End Sub


Sub queryForm(actionURL,ValueName,Value,optValueName,optValue,finalFilter,RecordCount,PageNoCount,PageSizeString,PageSize,PageNo)
'======================================================
'actionURL=""		'FORM提交的URL文件名
'ValueName="In_TeachType_Id,In_Plan_Year,In_Plan_Term"	'保留的变量名
'Value=TeachType_Id&","&Plan_Year&","&Plan_Term			'保留的变量值
'optValueName="s_Lesson_Name,s_TeacherName,d_Start_Date,d_End_Date"		'查找条件字段名
'optValue="课程名称,任课教师,开始时间,截止时间"						'查找条件名
'finalFilter=Request.Form("finalFilter")								'已经生成的条件
'RecordCount=result													'记录总数
'PageNoCount=rs.PageCount											'页总数
'PageSizeString="30,50,100"											'分页选项
'PageSize=rs.PageSize												'当前每页值
'PageNo=rs.AbsolutePage												'当前第几页
'======================================================
	Dim i
	Response.write "<script language=""javascript"" src=""../scripts/query.js""></script>"&vbcrlf
	Response.write "<form id=""query"" action="""&actionURL&""" method=""post"" onsubmit=""return chkField()"">"&vbcrlf

	'----开始 传参变量
'	Response.write Value
	ValueName=split(ValueName,",")
	Value=split(Value,",")
	For i=0 To ubound(ValueName)
		Response.write "	<input type=""hidden"" name="""&ValueName(i)&""" value="""&Value(i)&""">"&vbcrlf
	Next
	'----结束 传参变量

	Response.write vbcrlf
	Response.write "	<table cellspacing=4 cellpadding=0 width=""750""><tr align=center><td>"&vbcrlf

	'----开始 显示查询条件
	Response.write "		<select name=""field"" onchange=""ReloadOperator()"">"&vbcrlf
	optValueName=split(optValueName,",")
	optValue=split(optValue,",")
	For i=0 To ubound(optValueName)
		Response.write "			<option name="""&optValueName(i)&""">"&optValue(i)&"</option>"&vbcrlf
	Next
	Response.write "		</select>"&vbcrlf
	Response.write "		<select name=""operator"">"&vbcrlf
	Response.write "			<script>ReloadOperator()</script>"&vbcrlf
	Response.write "		</select>"&vbcrlf
	'----结束 显示查询条件

	Response.write "		<input type=""text"" name=""filter"" size=""10""  onkeypress=""checkKey()"">"&vbcrlf
	Response.write "		<input type=""hidden"" name=""finalFilter"" value="""&finalFilter&""">"&vbcrlf
	Response.write "		<input type=""submit"" value=""查找"" onclick=""genFilter()"">"&vbcrlf
	Response.write "		<input type=""submit"" value=""在结果中查找"" onclick=""genFinalFilter()"">"&vbcrlf

	'----开始 分页条件
	Response.write "		&nbsp;每页"&vbcrlf
	Response.write "		<select name=""pageSize"" onchange=""this.form.submit()"">"&vbcrlf
	PageSizeString=split(PageSizeString,",")
	For i=0 To ubound(PageSizeString)
		Response.write "			<option value="""&PageSizeString(i)&""""
		If PageSizeString(i)=PageSize Then Response.write "selected"
		Response.write ">"&PageSizeString(i)&"</option>"&vbcrlf
	Next
	Response.write "		</select>"&vbcrlf
	Response.write "		条&nbsp;转到"&vbcrlf

	Response.write "		<select name=""pageNo"" onchange=""this.form.submit()"">"&vbcrlf
	For i=1 To PageNoCount
		Response.write "			<option value="""&i&""""
		If i=Cint(PageNo) Then Response.write "selected"
		Response.write ">"&i&"</option>"&vbcrlf
	Next
	Response.write "		</select>"&vbcrlf
	Response.write "		页&nbsp;共"&RecordCount&"条"&vbcrlf
	'----结束 分页条件

	Response.write "	</td></tr></table>"&vbcrlf
	Response.write "</form>"&vbcrlf
End Sub


Sub SafeRequest(SqlStr, StrType)
'**************************************************************
'Author:华工ASP开发小组
'Name:判断传入参数里出成的字符次数过程(主要用于判断是否有注入漏洞)
'Created Date:	2005/4/24
'Purpose:通过传惨 ValueName,Value,ReturnURL生成HTML 后自动返回
'Relate pages:多用于保存、修改、删除后的页面调用
'Modify By:	Dajing
'Modify date:	2005/4/24

'SqlStr 要判断的字符串  
'StrType 判断出现字符的类型(每类型都对应的字符,字符存在过程前端的变量Strings中) 
'	例：传入变量SqlStr内容为:"select field from table union select filed2 from table2"
'		传入变量StrType内容为"RS"那么 select出现的次数装不能超过2 否则报错
'**************************************************************

	Dim Strings,Num,i
	Strings=";,0x,select,exec"
	Strings=split(Strings,",")
	StrType=ucase(StrType)
	If StrType="PROC" Then
		Num="1,1,1,2"				'用于一个过程语句 不能出现SELECT 
	ElseIf StrType="RS" Then
		Num="1,1,2,1"				'用于一个SELECT SQL语句
	ElseIf StrType="PARAM" Then
		Num="1,1,1,1"				'用于传参
	Else
		Num="1,1,10,1"				'用于多SELECT  SQL语句
	End If
	Num=split(Num,",")
	For i=0 to ubound(Strings)
		If Cint(RegExpTest(Strings(i),sqlStr))>=Cint(Num(i)) Then 
			Response.write "</select>"
			Response.write "<SCRIPT LANGUAGE='JavaScript'>"
			Response.write "alert('非法操作');"
			Response.write "window.history.back();"
			Response.write "</script>"
			Response.end
		End If 
	Next
End Sub

Sub GetRecordSetSafe(conn,rsNoLock,sqlStr,result)
	Set rsNoLock=Server.CreateObject("ADODB.RECORDSET")
	If IsEmpty(conn) Then Connect conn
	rsNoLock.activeConnection=conn
	rsNoLock.source=sqlStr
	rsNoLock.Open , ,AdOpenKeyset,AdLockReadOnly
	result=rsNoLock.RecordCount
End Sub

Function WriteLog(content)	' 添加日志记录
	Dim logfile,fso,stream,msg
	logfile=Server.MapPath("/log/"&toDateTime(Date,1)&".log")
	Set fso=Server.CreateObject("Scripting.FileSystemObject")
	Set stream=fso.OpenTextFile(logfile,8,true)
	msg="["&Time&"|"&Request.ServerVariables("REMOTE_ADDR")&"]"&content
	stream.WriteLine msg
	stream.Close
	Set fso=Nothing
End Function
Function WriteLogForReviewSystem(content)	' 添加论文电子评阅系统日志记录
	Dim logfile,fso,stream,msg
	logfile=Server.MapPath("/log/review/"&toDateTime(Date,1)&".log")
	Set fso=Server.CreateObject("Scripting.FileSystemObject")
	Set stream=fso.OpenTextFile(logfile,8,true)
	msg="["&Time&"|"&Request.ServerVariables("REMOTE_ADDR")&"]"&content
	stream.WriteLine msg
	stream.Close
	Set fso=Nothing
End Function
Function WriteLogForTutorSystem(content)	' 添加选导师系统日志记录
	Dim logfile,fso,stream,msg
	logfile=Server.MapPath("/log/tutor/"&toDateTime(Date,1)&".log")
	Set fso=Server.CreateObject("Scripting.FileSystemObject")
	Set stream=fso.OpenTextFile(logfile,8,true)
	msg="["&Time&"|"&Request.ServerVariables("REMOTE_ADDR")&"]"&content
	stream.WriteLine msg
	stream.Close
	Set fso=Nothing
End Function

Function getCurrentSemester()
	' 返回当前学期信息
	Dim start_year,cur_semester,semester_name,period_id
	Dim arr(3)
	If Month(Now)>=9 Or Month(Now)=1 Then
		If Month(Now)>=9 Then
			start_year=Year(Now)
		Else
			start_year=Year(Now)-1
		End If
		cur_semester=1
		semester_name="上"
	Else
		start_year=Year(Now)-1
		cur_semester=2
		semester_name="下"
	End If
	period_id=start_year&cur_semester
	arr(0)=start_year
	arr(1)=cur_semester
	arr(2)=semester_name
	arr(3)=period_id
	getCurrentSemester=arr
End Function
'========================
Function toSqlString(ByVal s)
	If IsNull(s) Then
		toSqlString="NULL"
		Exit Function
	End If
	s=Replace(Replace(Replace(Replace(s,"'","''"),vbCrLf,"'+CHAR(13)+CHAR(10)+'"),vbCr,"'+CHAR(13)+'"),vbLf,"'+CHAR(10)+'")
	If Len(s)=0 Then
		s="NULL"
	Else
		s="N'"&s&"'"
	End If
	toSqlString=s
End Function
Function toPlainString(ByVal s)
	If IsNull(s) Then Exit Function
	s=Replace(s,"""","&quot;")
	s=Replace(s,"<","&lt;")
	s=Replace(s,">","&gt;")
	s=Replace(s,vbNewLine,"<br/>")
	s=Replace(s,vbCr,"<br/>")
	s=Replace(s,vbLf,"<br/>")
	toPlainString=s
End Function
Function toJsString(ByVal s)
	If IsNull(s) Then Exit Function
	s=Replace(s,"""","\""")
	s=Replace(s,"'","\'")
	s=Replace(s,vbNewLine,"\n")
	toJsString=s
End Function
Function toPlainText(ByVal s)
	Dim r,arr,match
	Set r=New RegExp
	r.Pattern="<[^>]*>|&[A-z]+;"
	r.IgnoreCase=True
	r.Global=True
	Set arr=r.Execute(s)
	For Each match In arr
		s=Replace(s,match.Value,vbNullString)
	Next
	toPlainText=s
End Function
Function toEscapeString(ByVal s)
	Dim i,high,low,char,ret
	For i=1 To Len(s)
		char=Hex(AscW(Mid(s,i,1)))
		char=String(4-Len(char),"0")&char
		low=Right(char,2)
		high=Left(char,2)
		ret=ret&"%"&low&"%"&high
	Next
	toEscapeString=ret
End Function
Function toNumber(ByVal s)
	If IsNull(s) Then
		toNumber="0"
		Exit Function
	End If
	If Left(s,1)="." Then
		s="0"&s
	End If
	toNumber=s
End Function
Function toDateTime(d,fmt)
	If IsNull(d) Then
		toDateTime="0"
		Exit Function
	End If
	If d="0" Then
		toDateTime="未定"
	Else
		toDateTime=FormatDateTime(d,fmt)
	End If
End Function
Function toDataSizeString(value)
	Dim unit:unit=Array("字节","kB","MB","GB")
	Dim d,i:i=0
	d=value
	Do While d>=1024
		d=d/1024
		i=i+1
	Loop
	toDataSizeString=Round(d,2)&" "&unit(i)
End Function
%><!--#include file="mail.asp"-->