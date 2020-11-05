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
Const rsDetectPaperUploaded=1
Const rsRefusedDetect=2
Const rsAgreedDetect=3
Const rsDetectUnpassed=4
Const rsRedetectUnpassed=5
Const rsAgreedReview=6
Const rsMatchedReviewer=7
Const rsReviewed=8
Const rsReviewEval=9
Const rsDefencePaperUploaded=10
Const rsRefusedDefence=11
Const rsAgreedDefence=12
Const rsDefenceEval=13
Const rsInstructReviewPaperUploaded=14
Const rsRefusedInstructReview=15
Const rsAgreedInstructReview=16
Const rsInstructReviewPaperDetected=17	' 停用
Const rsMatchedInstructMember=18
Const rsInstructEval=19
Const rsFinalPaperUploaded=20

' 环节 Id
Const sectionUploadKtbgb = 1
Const sectionUploadZqkhb = 2
Const sectionUploadYdbyjs = 3
Const sectionUploadSpclb = 4
Const sectionUploadDetectReview = 5
Const sectionUploadReview = 6
Const sectionUploadDefence = 7
Const sectionUploadInstructReview = 8
Const sectionUploadFinal = 9
Const sectionAudit = 10
Const sectionInstructReview = 11
Const sectionReview = 12

' 审核意见类型
Const auditTypeKtbgb = 1
Const auditTypeZqkhb = 2
Const auditTypeYdbyjs = 3
Const auditTypeSpclb = 4
Const auditTypeDetectReview = 5
Const auditTypeReviewApp = 6
Const auditTypeReviewResult = 7
Const auditTypeDefence = 8
Const auditTypeInstructReviewDetect = 9
Const auditTypeInstructReview = 10

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
							'script', basePath(), arr[j], query_string));
					}
				} else {
					Response.Write(Format('<{0} type="text/javascript" src="{1}scripts/{2}.js{3}"></{0}>',
						'script', basePath(), value, query_string));
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
							basePath(), arr[j], query_string));
					}
				} else {
					Response.Write(Format('<link href="{0}css/{1}.css{2}" rel="stylesheet" type="text/css" />',
							basePath(), value, query_string));
				}
			}
		}
	}

	// 返回规范化路径名
	function resolvePath() {
		var arr=[];
		var delimeter;
		for(var i=0;i<arguments.length;++i) {
			arr.push(arguments[i]);
		}
		delimeter=arr[0].indexOf("/") != -1 ? "/" : "\\";
		return arr.join(delimeter).replace(/[\/\\]{2,}/g,delimeter);
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

Function basePath()
	basePath = "/PaperReview/"
End Function

Function baseLogPath()
	baseLogPath = "/log/PaperReview/"
End Function

Function tempPath()
	tempPath = resolvePath(basePath(), "temp/")
End Function

Function uploadBasePath(user_type, file_type)
	Dim client_path:client_path = Array("admin","student","tutor","expert")(user_type)
	uploadBasePath = basePath() & client_path & "/upload/" & file_type & "/"
End Function

Function exportBasePath(user_type)
	Dim client_path:client_path = Array("admin","student","tutor","expert")(user_type)
	exportBasePath = basePath() & client_path & "/export/"
End Function

Function reviewFileExportPath()
	reviewFileExportPath = resolvePath(basePath(),"expert/export/")
End Function

Function reviewFileVersionPath(filename)
	Dim basePath:basePath = reviewFileExportPath()
	reviewFileVersionPath=Array(Server.MapPath(resolvePath(basePath,filename&".pdf")),_
		Server.MapPath(resolvePath(basePath,filename&"_nostu.pdf")),_
		Server.MapPath(resolvePath(basePath,filename&"_noexp.pdf")))
End Function

' 若指定目录不存在，则新建目录
Function ensurePathExists(path)
	Dim fso: Set fso = CreateFSO()
	If Not fso.FolderExists(path) Then
		fso.CreateFolder(path)
	End If
	Set fso = Nothing
	ensurePathExists = True
End Function

' 返回与当前时间相关的字符序列
Function timestamp()
	timestamp=toDateTime(Now(),1)&Int(Timer)
End Function

Function serverVar(name)
	serverVar=Request.ServerVariables(name)
End Function

Function currentPage()
	currentPage=LCase(Mid(serverVar("SCRIPT_NAME"),Len(basePath())+1))
End Function

Function currentClient()
	Dim str:str=currentPage()
	Dim i:i=InStr(str,"/")
	If i=0 Then
		currentClient="none"
	Else
		currentClient=LCase(Left(str,InStr(str,"/")-1))
	End If
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
	ConnectDb conn
	sql="SELECT * FROM Configs"
	GetRecordSetNoLock conn,rs,sql,count
	If rs.EOF Then
		CloseRs rs
		CloseConn conn
		Response.Redirect basePath()&"error.asp?no-config"
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
	ConnectDb conn
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
	ConnectDb conn
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
		Set getSectionInfo=Nothing
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
	Dim conn:ConnectDb conn
	Dim sql:sql="SELECT * FROM ViewActivities WHERE Id=?"
	Dim ret:Set ret=ExecQuery(conn,sql,CmdParam("Id",adInteger,4,activity_id))
    Dim rs:Set rs=ret("rs")
    Dim count:count=ret("count")
	If count=0 Then
		getActivityInfo=Null
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
		stu_type_id=stutypeME Or stutypeMBA Or stutypeMPAcc ' Or stutypeEMBA
	End If
	Dim conn:ConnectDb conn
	Dim sql:sql="SELECT TOP 1 Id,Name,SemesterId,SemesterName FROM ViewActivities WHERE Valid=1 AND (?&StuType)<>0 ORDER BY SemesterId DESC"
	Dim ret:Set ret=ExecQuery(conn,sql,CmdParam("StuTypes",adInteger,4,stu_type_id))
    Dim rs:Set rs=ret("rs")
    Dim count:count=ret("count")
	If count=0 Then
		Set getLastActivityInfoOfStuType=Nothing
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

' 根据姓名获取教师 ID
Function getTeacherIdByName(ByVal name)
	If IsNull(name) Then
		getTeacherIdByName=-1
		Exit Function
	End If
	Dim conn,rs,sql,count
	name=Replace(name," ",vbNullString)
	name=Replace(name,"　",vbNullString)
	ConnectJWDb conn
	sql="SELECT TEACHERID,TEACHERNAME FROM TEACHER_INFO WHERE TEACHERNAME="&toSqlString(name)&" AND VALID=0"
	GetRecordSetNoLock conn,rs,sql,count
	If rs.EOF Then
		getTeacherIdByName=-1
	Else
		getTeacherIdByName=rs("TEACHERID")
	End If
	CloseRs rs
	CloseConn conn
End Function

' 根据姓名获取评阅专家 ID
Function getExpertIdByName(ByVal name)
	If IsNull(name) Then
		getExpertIdByName=-1
		Exit Function
	End If
	Dim conn,rs,sql
	ConnectDb conn
	sql="SELECT TEACHER_ID,EXPERT_NAME FROM Experts WHERE EXPERT_NAME="&toSqlString(name)
	Set rs=conn.Execute(sql)
	If rs.EOF Then
		getExpertIdByName=-1
	Else
		getExpertIdByName=rs("TEACHER_ID")
	End If
	CloseRs rs
	CloseConn conn
End Function

' 根据联系号码获取评阅专家 ID
Function getExpertIdByTelephone(ByVal telephone)
	If IsNull(telephone) Then
		getExpertIdByTelephone=-1
		Exit Function
	End If
	Dim conn,rs,sql
	ConnectDb conn
	sql="SELECT TEACHER_ID,EXPERT_NAME FROM Experts WHERE "&toSqlString(telephone)&" IN (TELEPHONE, MOBILE)"
	Set rs=conn.Execute(sql)
	If rs.EOF Then
		getExpertIdByTelephone=-1
	Else
		getExpertIdByTelephone=rs("TEACHER_ID")
	End If
	CloseRs rs
	CloseConn conn
End Function

' 获取审核记录信息
Function getAuditInfo(paper_id,audit_file,audit_type)
	Dim conn:ConnectDb conn
	Dim sql,ret
	If Not IsNull(audit_file) And Not IsNull(audit_type) Then
		sql="SELECT * FROM ViewAuditRecords WHERE DissertationId=? AND AuditFile=? AND AuditType=? ORDER BY AuditTime DESC"
		Set ret=ExecQuery(conn,sql,_
			CmdParam("paper_id",adInteger,4,paper_id),_
			CmdParam("audit_file",adVarWChar,50,audit_file),_
			CmdParam("audit_type",adInteger,4,audit_type))
	ElseIf Not IsNull(audit_type) Then
		sql="SELECT * FROM ViewAuditRecords WHERE DissertationId=? AND AuditType=? ORDER BY AuditTime DESC"
		Set ret=ExecQuery(conn,sql,_
			CmdParam("paper_id",adInteger,4,paper_id),_
			CmdParam("audit_type",adInteger,4,audit_type))
	Else
		sql="SELECT * FROM ViewAuditRecords WHERE DissertationId=? ORDER BY AuditTime DESC"
		Set ret=ExecQuery(conn,sql,CmdParam("paper_id",adInteger,4,paper_id))
	End If
    Dim rs:Set rs=ret("rs")
    Dim count:count=ret("count")
	Dim result:ReDim result(count-1)
	Dim dict
	If count=0 Then
		Set dict = CreateDictionary()
		dict.Add "Id","0"
		ReDim result(0):Set result(0) = dict
		getAuditInfo = result
		CloseRs rs
		CloseConn conn
		Exit Function
	End If
	Dim i
	For i=0 To UBound(result)
		Set dict = CreateDictionary()
		dict.Add "Id", rs("Id").Value
		dict.Add "DissertationId",rs("DissertationId").Value
		dict.Add "AuditFile",rs("AuditFile").Value
		dict.Add "AuditType",rs("AuditType").Value
		dict.Add "AuditTypeName",rs("AuditTypeName").Value
		dict.Add "AuditTime",rs("AuditTime").Value
		dict.Add "AuditorName",rs("AuditorName").Value
		dict.Add "IsPassed",rs("IsPassed").Value
		dict.Add "Comment",rs("Comment").Value
		Set result(i) = dict
		rs.MoveNext()
	Next
	getAuditInfo = result
	CloseRs rs
	CloseConn conn
End Function

' 获取指定论文和审核类型的最新审核记录信息
Function getLatestAuditInfo(dissertationId,auditType)
	Dim conn,sql,ret,rs
	ConnectDb conn
	sql="EXEC spGetAuditRecord ?,?,NULL"
	Set ret=ExecQuery(conn,sql,_
		CmdParam("DissertationId",adInteger,4,dissertationId),_
		CmdParam("AuditType",adInteger,4,auditType))
	Set rs=ret("rs")
	Dim dict:Set dict=CreateDictionary()
	If rs.EOF Then
		dict.Add "Id","0"
		Set getLatestAuditInfo = dict
		CloseRs rs
		CloseConn conn
		Exit Function
	End If
	dict.Add "Id", rs("Id").Value
	dict.Add "DissertationId", rs("DissertationId").Value
	dict.Add "AuditTypeName", rs("AuditTypeName").Value
	dict.Add "AuditTime",rs("AuditTime").Value
	dict.Add "AuditorName", rs("AuditorName").Value
	dict.Add "IsPassed",rs("IsPassed").Value
	dict.Add "Comment", rs("Comment").Value
	Set getLatestAuditInfo = dict
	CloseRs rs
	CloseConn conn
End Function

Function ensureSystemOpen()
	Dim configs:Set configs=getSystemConfigs()
	If configs("Status")="closed" Then
		Response.Redirect basePath()&"error.asp?closed"
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
Dim dictSystemOptions:Set dictSystemOptions = CreateDictionary()
'dictSystemOptions.Add "send_accept_email",			Array("导师确认填报时，向学生发送通知邮件", 0)

// 获取系统选项值（没有则使用默认值）
Function getSystemOption(name,stu_type_id)
	If Not dictSystemOptions.Exists(name) Then Exit Function
	Dim conn,rs,sql,count
	ConnectDb conn
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

Function CreateFSO()
	Set CreateFSO = Server.CreateObject("Scripting.FileSystemObject")
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
dictScripts.Add "paper",			"paper"
dictScripts.Add "tutor",			"tutor"
dictScripts.Add "upload",			"upload"
dictScripts.Add "fillInTable",		"fillInTable"
dictScripts.Add "uploadPaper",		"uploadPaper"
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
	If Len(viewStateData) = 0 Then viewStateData=Null
	Dim conn,sql
	ConnectDb conn
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
	ConnectDb conn
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

' 生成日期格式文件名
Function generateDateTimeFilename(ext)
	generateDateTimeFilename = FormatDateTime(Now(),1)&Int(Timer)&"."&ext
End Function

' 添加日志记录
Function writeLog(content)
	Dim logfile,fso,stream,msg
	logfile=Server.MapPath(baseLogPath()&toDateTime(Date,1)&".log")
	Set fso=CreateFSO()
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

' 显示评阅书评价指标信息
Function loadReviewScoringInfo(review_type,formcode,power1code,power2code)
	Dim stream,filepath
	Dim data,numScores1,numScores2,nameScore1,nameScore2,power1,power2,power2sum,detail
	Dim c,i,j,k,tmp,tmp2,tmp3
	Dim retcode,retpower1,retpower2
	Set stream=Server.CreateObject("ADODB.Stream")
	filepath=Server.MapPath(basePath()&"expert/data/scoreform"&review_type&".html")
	stream.Mode=3
	stream.Type=2
	stream.Charset="utf-8"
	stream.Open()
	stream.LoadFromFile filepath
	data=Split(stream.ReadText(),"|")
	numScores1=data(2)
	c=3
	For i=1 To numScores1
		power2sum=0
		nameScore1=data(c)
		power1=data(c+1)
		numScores2=data(c+2)
		c=c+3
		tmp=vbNullString
		If Len(retpower1) Then retpower1=retpower1&","
		retpower1=retpower1&power1
		If i>1 Then
			retpower2=retpower2&",["
		Else
			retpower2=retpower2&"["
		End If
		For j=1 To numScores2
			nameScore2=Replace(data(c),vbNewLine,"<br/>")
			detail=Split(data(c+1),"；")
			power2=data(c+2)
			If IsNumeric(power2) Then
				power2=power2*1
				power2sum=power2sum+power2
			Else
				power2=0
			End If
			If j>1 Then retpower2=retpower2&","
			retpower2=retpower2&power2
			tmp3=vbNullString
			For k=0 To UBound(detail)
				tmp3=tmp3&"<li>"&detail(k)&"</li>"
			Next
			If j=1 Then
				tmp2=data(0)
			Else
				tmp2=data(1)
			End If
			tmp2=Replace(tmp2,"$numScores2",numScores2)
			tmp2=Replace(tmp2,"<score2 />",i&"."&j&"&nbsp;"&nameScore2)
			tmp2=Replace(tmp2,"<detail />",tmp3)
			tmp2=Replace(tmp2,"<power2 />",power2*100)
			tmp=tmp&tmp2
			c=c+3
		Next
		retpower2=retpower2&"]"
		tmp=Replace(tmp,"<score1 />",nameScore1)
		tmp=Replace(tmp,"<power1 />",power1*100)
		tmp=Replace(tmp,"<power2sum />",power2sum*100)
		retcode=retcode&tmp
	Next
	retpower1="["&retpower1&"]"
	retpower2="["&retpower2&"]"
	stream.Close()
	Set stream=Nothing
	formcode=retcode
	power1code=retpower1
	power2code=retpower2
End Function

' 解压缩 RAR/ZIP 文件
Function extractFile(file, destPath)
	Dim exe, cmd, fileType
	fileType = LCase(Mid(file, InStrRev(file, ".")))
	If fileType = ".zip" Then
		exe = Server.MapPath("rar/unzip.exe")
		cmd = Format("""{0}"" -j ""{1}"" -d ""{2}""",exe,file,destPath)
	ElseIf fileType = ".rar" Then
		exe = Server.MapPath("rar/Rar.exe")
		cmd = Format("""{0}"" e -o+ ""{1}"" ""{2}""",exe,file,destPath)
	End If
	Dim wsh:Set wsh=Server.CreateObject("WScript.Shell")
	Dim exec:Set exec=wsh.Exec(cmd)
	extractFile = exec.StdOut.ReadAll()
	Set exec=Nothing
	Set wsh=Nothing
End Function

Session("DebugInfo")=Empty
Sub debug(info)
	Session("DebugInfo")=Session("DebugInfo")&vbNewLine&info
End Sub

If currentPage()<>"default.asp" And IsEmpty(Session("id")) Then
	ensureSystemOpen()
End If

Public dictStuTypes:Set dictStuTypes=CreateDictionary()
With dictStuTypes
	.Add 5,Array("工程硕士", "ME")
	.Add 6,Array("工商管理硕士", "MBA")
	.Add 7,Array("高级管理人员工商管理硕士", "EMBA")
	.Add 9,Array("会计硕士", "MPAcc")
End With

Public arrDiplomaName:arrDiplomaName=Array("","专科","本科","硕士研究生","博士研究生","博士后")
Public arrDefenceResult:arrDefenceResult=Array("未录入","优秀","良好","一般","较差")
Public arrGrantDegreeResult:arrGrantDegreeResult=Array("未录入","同意毕业及授予硕士学位","同意毕业，但不同意授予硕士学位","不同意毕业")
Public arrReviewFileStat:arrReviewFileStat=Array("不开放显示","仅向导师显示","完全开放显示")
Public arrDefaultFileListName:arrDefaultFileListName=Array("","开题报告表","开题论文","中期考核表","中期论文","预答辩意见书","预答辩论文","答辩及授予学位审批材料","送检论文","送审论文","答辩论文","教指委盲评论文","定稿论文","送检论文检测报告","教指委盲评论文检测报告","硕士学位论文送审申请表","论文评阅书（专家一）","论文评阅书（专家二）","论文评阅书")
Public arrDefaultFileListNamePostfix:arrDefaultFileListNamePostfix=Array("","开题报告表","开题论文","中期考核表","中期论文","预答辩意见书","预答辩论文","答辩审批材料","","","","","","送检论文检测报告","教指委盲评论文检测报告","送审申请表","论文评阅书（专家一）","论文评阅书（专家二）","论文评阅书")
Public arrDefaultFileListPath:arrDefaultFileListPath=Array("","student/upload","student/upload","student/upload","student/upload","student/upload","student/upload","student/upload","student/upload","student/upload","student/upload","student/upload","student/upload","admin/upload/detect_report","admin/upload/detect_report","tutor/export","expert/export","expert/export","expert/export")
Public arrDefaultFileListField:arrDefaultFileListField=Array("","TABLE_FILE1","TBL_THESIS_FILE1","TABLE_FILE2","TBL_THESIS_FILE2","TABLE_FILE3","TBL_THESIS_FILE3","TABLE_FILE4","THESIS_FILE","THESIS_FILE2","THESIS_FILE3","THESIS_FILE4","THESIS_FILE5","DETECT_REPORT","INSTRUCT_REVIEW_DETECT_REPORT","REVIEW_APP","ReviewFile1","ReviewFile2","")
' 禁止缓存，仅在需要客户端重新请求 JS/CSS/HTML 等静态文件时使用
Response.CacheControl="no-cache"
Response.Charset="utf-8"
Response.Buffer = True
%>
<!--#include file="JSONWriter.inc"-->
<!--#include file="PinyinQuery.inc"-->
<!--#include file="conversion.asp"-->
<!--#include file="database.asp"-->
<!--#include file="mail.asp"-->
<!--#include file="helper.asp"-->
<!DOCTYPE html>