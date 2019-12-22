<%
'=========================================
' 全局成员定义文件
' 2019-05-10
'=========================================

' 用户类型
Const usertypeAdmin=0
Const usertypeStudent=1
Const usertypeTutor=2
Const usertypeExpert=3

' 通知类型
Const notifytypeMail=0
Const notifytypeSms=1

' 学生类型
Const stutypeME=16
Const stutypeMBA=32
Const stutypeEMBA=64
Const stutypeMPAcc=256

' 表格审核状态
Const tpNone=0
Const tpTbl1Uploaded=1
Const tpTbl1Unpassed=2
Const tpTbl1Passed=3
Const tpTbl2Uploaded=4
Const tpTbl2Unpassed=5
Const tpTbl2Passed=6
Const tpTbl3Uploaded=7
Const tpTbl3Unpassed=8
Const tpTbl3Passed=9
Const tpTbl4Uploaded=10
Const tpTbl4Unpassed=11
Const tpTbl4Passed=12

' 论文状态
Const rsNone=0
Const rsDetectThesisUploaded=1
Const rsNotAgreeDetect=2
Const rsAgreeDetect=3
Const rsDetectUnpassed=4
Const rsRedetectUnpassed=5
Const rsRedetectPassed=6
Const rsNotAgreeReview=7
Const rsAgreeReview=8
Const rsMatchExpert=9
Const rsReviewed=10
Const rsReviewEval=11
Const rsModifyThesisUploaded=12
Const rsModifyUnpassed=13
Const rsModifyPassed=14
Const rsDefenceEval=15
Const rsInstructEval=16
Const rsFinalThesisUploaded=17

' 环节 Id
Const sectionUploadKtbg = 1
Const sectionUploadZqjcb = 2
Const sectionUploadYdbyjs = 3
Const sectionUploadSpcl = 4
Const sectionUploadDetectReview = 5
Const sectionUploadReview = 6
Const sectionUploadDefence = 7
Const sectionUploadFinal = 8
Const sectionAudit = 9
Const sectionReview = 10

%>
<script language="jscript" runat="server">
	// 获取指定表单参数
	function getFormParams() {
		var dict=CreateDictionary();
		for (i=0;i<arguments.length;++i) {
			dict.Add(arguments[i],Request.Form(arguments[i]));
		}
		return dict;
	}

	// 引用指定脚本
	function useScript() {
		for(var i=0;i<arguments.length;++i) {
			var alias=arguments[i];
			var query_string;
			if (alias.charAt(0)==='*') {
				query_string='?t='+parseInt(Math.random()*1e5);
				alias=alias.substr(1);
			} else {
				query_string='';
			}
			if (!dictScripts.Exists(alias)) {
				Response.Write(Format('<{0} type="text/javascript">console.warn("useScript(): 脚本\\"{1}\\"未注册。");</{0}>',
					'script', alias));
			} else {
				var value=dictScripts.Item(alias);
				if (value.indexOf("|")!==-1) {
					var arr=value.split("|");
					for(var j=0;j<arr.length;++j) {
						Response.Write(Format('<{0} type="text/javascript" src="{1}scripts/{2}.js{3}"></{0}>',
							'script', baseUrl(), arr[j], query_string));
					}
				} else {
					Response.Write(Format('<{0} type="text/javascript" src="{1}scripts/{2}.js{3}"></{0}>',
						'script', baseUrl(), value, query_string));
				}
			}
		}
	}

	// 引用指定样式表
	function useStylesheet() {
		for(var i=0;i<arguments.length;++i) {
			var alias=arguments[i];
			var query_string;
			if (alias.charAt(0)==='*') {
				query_string='?t='+parseInt(Math.random()*1e5);
				alias=alias.substr(1);
			} else {
				query_string='';
			}
			if (!dictStylesheets.Exists(alias)) {
				Response.Write(Format('<{0} type="text/javascript">console.warn("useStylesheet(): 样式表\\"{0}\\"未注册。");</{0}>',
					'script', alias));
			} else {
				var value=dictStylesheets.Item(alias);
				if (value.indexOf("|")!==-1) {
					var arr=value.split("|");
					for(var j=0;j<arr.length;++j) {
						Response.Write(Format('<link href="{0}css/{1}.css{2}" rel="stylesheet" type="text/css" />',
							baseUrl(), arr[j], query_string));
					}
				} else {
					Response.Write(Format('<link href="{0}css/{1}.css{2}" rel="stylesheet" type="text/css" />',
							baseUrl(), value, query_string));
				}
			}
		}
	}
</script><%

Function hasPrivilege(pArr,privilege)
	If IsEmpty(pArr) Then
		hasPrivilege=False
		Exit Function
	End If
	pArr=split(pArr,",")
	For i=0 To Ubound(pArr)-1
		If pArr(i)=privilege Then
			hasPrivilege=True
			Exit Function
		End If
	Next
	hasPrivilege=False
End Function

Function baseUrl()
	baseUrl = "/ThesisReview/"
End Function

Function serverVar(name)
	serverVar=Request.ServerVariables(name)
End Function

Function currentPage()
	currentPage=LCase(Mid(serverVar("SCRIPT_NAME"),Len(baseUrl())+1))
End Function

Function currentClient()
	Dim str:str=currentPage()
	currentClient=LCase(Left(str,InStr(str,"/")-1))
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
	period_id=Int(start_year&cur_semester)
	arr(0)=start_year
	arr(1)=cur_semester
	arr(2)=semester_name
	arr(3)=period_id
	getCurrentSemester=arr
End Function

Function getSystemConfigs()
	Dim conn,rs,sql,count
	Connect conn
	sql="SELECT * FROM Configs"
	GetRecordSetNoLock conn,rs,sql,count
	If rs.EOF Then
		CloseRs rs
		CloseConn conn
		Response.Redirect baseUrl()&"error.asp?no-config"
		Exit Function
	End If

	Dim configs:Set configs=CreateDictionary()
	Dim i
	For i=0 To rs.Fields.Count-1
		configs.Add rs.Fields(i).Name, rs(i).Value
	Next
	CloseRs rs
	CloseConn conn
	Set getSystemConfigs=configs
End Function

' 获取指定评阅活动的开放状态
Function isActivityOpen(activity_id)
	Dim conn,rs,sql
	Connect conn
	sql="SELECT IsOpen FROM Activities WHERE Id=?"
	Dim ret:Set ret=ExecQuery(conn,sql,CmdParam("Id",adInteger,4,activity_id))
	Set rs=ret("rs")
	If rs.EOF Then
		isActivityOpen=False
	Else
		isActivityOpen=rs(0).Value
	End If
	CloseRs rs
	CloseConn conn
End Function

' 获取环节模板信息，或指定评阅活动的环节信息
Function getSectionInfo(activity_id, stu_type_id, section_id)
	Dim conn,rs,sql,count,ret
	Connect conn
	If IsNull(activity_id) Then
		sql="SELECT Id,Name,ClientType,ClientTypeName FROM ViewSections WHERE ID=?"
		Set ret=ExecQuery(conn,sql,CmdParam("SectionId",adInteger,4,section_id))
	Else
		sql="SELECT SectionId Id,SectionName Name,StuType,StuTypeName,ClientType,ClientTypeName,Enabled,StartTime,EndTime FROM ViewActivityPeriods WHERE ActivityId=? AND StuType=? AND SectionId=?"
		Set ret=ExecQuery(conn,sql,_
			CmdParam("ActivityId",adInteger,4,activity_id),_
			CmdParam("StuType",adInteger,4,stu_type_id),_
			CmdParam("SectionId",adInteger,4,section_id))
	End If
	Set rs=ret("rs")
	count=ret("count")

	If rs.EOF Then
		getSectionInfo=Null
		CloseRs rs
		CloseConn conn
		Exit Function
	End If
	Dim dict: Set dict=CreateDictionary()
	Dim i
	For i=0 To rs.Fields.Count-1
		dict.Add rs.Fields(i).Name, rs(i).Value
	Next
	Set getSectionInfo=dict
	CloseRs rs
	CloseConn conn
End Function

' 获取指定的评阅活动信息
Function getActivityInfo(activity_id)
	If IsNull(activity_id) Then
		Exit Function
	End If
	Dim conn:Connect conn
	Dim sql:sql="SELECT * FROM ViewActivities WHERE Id=?"
	Dim ret:Set ret=ExecQuery(conn,sql,CmdParam("Id",adInteger,4,activity_id))
    Dim rs:Set rs=ret("rs")
    Dim count:count=ret("count")
	If count=0 Then
		getLastActivityInfoOfStuType=Null
		Exit Function
	End If
	Dim dict:Set dict=CreateDictionary()
	Dim i
	For i=0 To rs.Fields.Count-1
		dict.Add rs.Fields(i).Name, rs(i).Value
	Next
	CloseRs rs
	CloseConn conn
	Set getActivityInfo=dict
End Function

' 获取指定学生类型中最新的评阅活动信息
Function getLastActivityInfoOfStuType(stu_type_id)
	If IsNull(stu_type_id) Then
		stu_type_id=stutypeME Or stutypeMBA Or stutypeMPAcc Or stutypeEMBA
	End If
	Dim conn:Connect conn
	Dim sql:sql="SELECT TOP 1 Id,Name FROM ViewActivities WHERE Valid=1 AND (?&StuType)<>0 ORDER BY SemesterId DESC"
	Dim ret:Set ret=ExecQuery(conn,sql,CmdParam("StuTypes",adInteger,4,stu_type_id))
    Dim rs:Set rs=ret("rs")
    Dim count:count=ret("count")
	If count=0 Then
		getLastActivityInfoOfStuType=Null
		Exit Function
	End If
	Dim dict:Set dict=CreateDictionary()
	Dim i
	For i=0 To rs.Fields.Count-1
		dict.Add rs.Fields(i).Name, rs(i).Value
	Next
	CloseRs rs
	CloseConn conn
	Set getLastActivityInfoOfStuType=dict
End Function

' 比较指定时刻与某个时间范围的关系，0=在时间范围内，-1=在时间范围之前，1=在时间范围之后
Function compareTimeRange(time,time1,time2)
	If DateDiff("s", time1, time)<0 Then
		compareTimeRange=-1
	ElseIf DateDiff("s", time2, time)>0 Then
		compareTimeRange=1
	Else
		compareTimeRange=0
	End If
End Function

' 比较当前时刻与某个时间范围的关系
Function compareNowWithTimeRange(time1,time2)
	compareNowWithTimeRange=compareTimeRange(Now,time1,time2)
End Function

' 比较当前时间与指定环节的开放时间的关系
Function compareNowWithSectionTime(section)
	If Not section("Enabled") Then
		compareNowWithSectionTime=-2
		Exit Function
	End If
	compareNowWithSectionTime=compareNowWithTimeRange(section("StartTime"), section("EndTime"))
End Function

Function ensureSystemOpen()
	Dim configs:Set configs=getSystemConfigs()
	If configs("Status")="closed" Then
		Response.Redirect baseUrl()&"error.asp?closed"
		Exit Function
	End If
	ensureSystemOpen=1
End Function

Function verifyPasswordSecurityLevel(ByVal password)
	Dim ret
	If Len(password)<8 Then
		ret=0
	Else
		Dim str_types:str_types=0
		If isMatched("\d",password,False) Then str_types=str_types+1
		If isMatched("[A-Z]",password,False) Then str_types=str_types+1
		If isMatched("[a-z]",password,False) Then str_types=str_types+1
		If isMatched("[`~!@#$%\^&*()\-=_+\[\]{}\\|;':"",./<>?]",password,False) Then str_types=str_types+1
		If str_types<3 Then
			ret=1
		Else
			ret=2
		End If
	End If
	verifyPasswordSecurityLevel=ret
End Function
Dim arrSecurityLevelName:arrSecurityLevelName=Array("低强度","中等强度","高强度")

// 注册系统选项
Dim dictSystemOptions:Set dictSystemOptions = Server.CreateObject("Scripting.Dictionary")
'dictSystemOptions.Add "send_accept_email",			Array("导师确认填报时，向学生发送通知邮件", 0)

// 获取系统选项值（没有则使用默认值）
Function getSystemOption(name,stu_type_id)
	If Not dictSystemOptions.Exists(name) Then Exit Function
	Dim conn,rs,sql,count
	Connect conn
	sql="EXEC spGetSystemOption ?,?,?"
	Dim ret:Set ret=ExecQuery(conn,sql,_
		CmdParam("Name",adVarWChar,50,name),_
		CmdParam("StuTypeId",adInteger,4,stu_type_id),_
		CmdParam("DefaultValue",adInteger,4,dictSystemOptions.Item(name)(1)))
	Set rs=ret("rs")
	count=ret("count")
	getSystemOption=rs(0).Value
	CloseRs rs
	CloseConn conn
End Function

Function CreateDictionary()
	Dim ret: Set ret = Server.CreateObject("Scripting.Dictionary")
	Set CreateDictionary = ret
	Set ret = Nothing
End Function

// 注册脚本别名
Dim dictScripts:Set dictScripts=CreateDictionary()
dictScripts.Add "admin",			"admin"
dictScripts.Add "common",			"utils|query"
dictScripts.Add "drafting",			"drafting"
dictScripts.Add "expert",			"expert"
dictScripts.Add "expertList",		"expertList"
dictScripts.Add "jeasyui",			"jeasyui/jquery.easyui.min|jeasyui/easyloader"
dictScripts.Add "jquery",			"jquery-1.12.4.min"
dictScripts.Add "keywordList",		"keywordList"
dictScripts.Add "notifyList",		"notifyList"
dictScripts.Add "review",			"review"
dictScripts.Add "reviewSettings",	"reviewSettings"
dictScripts.Add "student",			"student"
dictScripts.Add "systemSettings",	"systemSettings"
dictScripts.Add "thesis",			"thesis"
dictScripts.Add "tutor",			"tutor"
dictScripts.Add "upload",			"upload"
dictScripts.Add "fillInTable",		"fillInTable"
dictScripts.Add "uploadThesis",		"uploadThesis"
dictScripts.Add "viewState",		"viewState"

// 注册样式表别名
Dim dictStylesheets:Set dictStylesheets=CreateDictionary()
dictStylesheets.Add "admin",		"admin"
dictStylesheets.Add "global",		"global"
dictStylesheets.Add "jeasyui",		"jeasyui/themes/icon|jeasyui/themes/bootstrap/easyui"
dictStylesheets.Add "student",		"student"
dictStylesheets.Add "tutor",		"tutor"

' 设置视图状态
Function setViewState(userId,userType,viewName,viewStateData)
	Dim conn,sql
	Connect conn
	sql="EXEC spSetViewState ?,?,?,?"
	ExecNonQuery conn,sql,_
		CmdParam("UserId",adInteger,4,userId),_
		CmdParam("UserType",adInteger,4,userType),_
		CmdParam("ViewName",adVarWChar,50,viewName),_
		CmdParam("ViewStateData",adVarWChar,65535,viewStateData)
	CloseConn conn
End Function

' 获取视图状态
Function getViewState(userId,userType,viewName)
	Dim conn,sql,ret,rs
	Connect conn
	sql="EXEC spGetViewState ?,?,?"
	Set ret=ExecQuery(conn,sql,_
		CmdParam("UserId",adInteger,4,userId),_
		CmdParam("UserType",adInteger,4,userType),_
		CmdParam("ViewName",adVarWChar,50,viewName))
	Set rs=ret("rs")
	If rs.EOF Then
		getViewState=Null
	Else
		getViewState=rs(0).Value
	End If
	CloseConn conn
End Function

' 添加日志记录
Function writeLog(content)
	Dim logfile,fso,stream,msg
	logfile=Server.MapPath("/log/ThesisReview/"&toDateTime(Date,1)&".log")
	Set fso=Server.CreateObject("Scripting.FileSystemObject")
	Set stream=fso.OpenTextFile(logfile,8,true)
	msg="["&Time&"]"&content
	stream.WriteLine msg
	stream.Close
	Set fso=Nothing
	writeLog=1
End Function

' 添加通知邮件（短信）发送事件记录
Function writeNotificationEventLog(user_type,user_name,trigger_operation,receive_user_type,receive_user_name,recipient,notify_type,is_sent)
	Dim user_type_names:user_type_names=Array("教务员","学生","导师","专家")
	Dim log_text:log_text=Format("{0}[{1}]{2}，发送通知{3}给{4}[{5}:{6}]{7}。",_
		user_type_names(user_type),user_name,trigger_operation,Array("邮件","短信")(notify_type),_
		user_type_names(receive_user_type),receive_user_name,recipient,_
		Array("失败","成功")(Abs(Int(is_sent))))
	writeNotificationEventLog=writeLog(log_text)
End Function

' 添加操作事件记录
Function writeEventLog(user_type,user_name,trigger_operation)
	Dim user_type_names:user_type_names=Array("教务员","学生","导师","专家")
	Dim log_text:log_text=Format("{0}[{1}]{2}。",user_type_names(user_type),user_name,trigger_operation)
	writeEventLog=writeLog(log_text)
End Function

' 显示错误提示
Function showErrorPage(content, title)
	Dim client_name, class_prop, statement
	client_name=currentClient()
	If client_name="expert" Then
		class_prop="class=""exp"""
		client_name="tutor"
	Else
		class_prop="bgcolor=""ghostwhite"""
	End If
	If isMatched("(expertProfile|fetchDocument|thesisDetail)\.asp",currentPage(),True) Then
		statement="closeWindow();"
	Else
		statement="history.go(-1);"
	End If
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="theme-color" content="#2D79B2" />
<title><%=title%></title>
<% useStylesheet client_name, "jeasyui" %>
<% useScript "jquery", "jeasyui", "common" %>
</head>
<body <%=class_prop%>>
<script type="text/javascript">
	$.messager.alert("<%=toJsString(title)%>", "<%=toJsString(content)%>", "error",
		function() {
			<%=statement%>
		}
	);
</script>
</body></html><%

	Response.End()
	showErrorPage=1
End Function

Session("DebugInfo")=Empty
Sub debug(info)
	Session("DebugInfo")=Session("DebugInfo")&vbNewLine&info
End Sub

If currentPage()<>"default.asp" And IsEmpty(Session("id")) Then
	ensureSystemOpen()
End If

Dim dictStuTypes:Set dictStuTypes=CreateDictionary()
With dictStuTypes
	.Add 5,Array("工程硕士", "ME")
	.Add 6,Array("工商管理硕士", "MBA")
	.Add 7,Array("高级管理人员工商管理硕士", "EMBA")
	.Add 9,Array("会计硕士", "MPAcc")
End With

Dim arrDiplomaName:arrDiplomaName=Array("","专科","本科","硕士研究生","博士研究生","博士后")
Dim arrDefenceResult:arrDefenceResult=Array("未录入","优秀","良好","一般","较差")
Dim arrGrantDegreeResult:arrGrantDegreeResult=Array("未录入","同意毕业及授予学位","同意毕业，但不同意授予学位","不同意毕业")
Dim arrReviewFileStat:arrReviewFileStat=Array("不开放显示","仅向导师显示","完全开放显示")
' 禁止缓存，仅在需要客户端重新请求 JS/CSS/HTML 等静态文件时使用
Response.CacheControl="no-cache"
Response.Charset="utf-8"
%>
<!--#include file="JSONWriter.inc"-->
<!--#include file="PinyinQuery.inc"-->
<!--#include file="conversion.asp"-->
<!--#include file="database.asp"-->
<!--#include file="mail.asp"-->
<!--#include file="helper.asp"-->