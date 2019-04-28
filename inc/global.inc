<%
'=========================================
' 专业学位论文电子评阅系统全局成员定义文件
' 2016-09-11
'=========================================

' TASK_PROGRESS
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
' REVIEW_STATUS
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

Const STUTYPE_COUNT=4
Const OPRTYPE_COUNT=8

Class StudentClientInfo
	Private opentime_start(),opentime_end()
	Private m_client_status()
	Private m_systemstatus
	Private Sub Class_Initialize()
		ReDim opentime_start(OPRTYPE_COUNT)
		ReDim opentime_end(OPRTYPE_COUNT)
		ReDim m_client_status(STUTYPE_COUNT,OPRTYPE_COUNT)
	End Sub
	Public Function getOpentime(Opr,TimeType)
		If TimeType=STUCLI_OPENTIME_START Then
			getOpentime=opentime_start(Opr)
		ElseIf TimeType=STUCLI_OPENTIME_END Then
			getOpentime=opentime_end(Opr)
		End If
	End Function
	Public Function setOpentime(Opr,TimeType,ByVal Value)
		If IsNull(Value) Then Value=0
		If TimeType=STUCLI_OPENTIME_START Then
			opentime_start(Opr)=Value
		ElseIf TimeType=STUCLI_OPENTIME_END Then
			opentime_end(Opr)=Value
		End If
		setOpentime=1
	End Function
	Public Function getClientStatus(StuType,Opr)
		Dim i
		Select Case StuType
		Case STUCLI_STUTYPE_ME:i=1
		Case STUCLI_STUTYPE_MBA:i=2
		Case STUCLI_STUTYPE_EMBA:i=3
		Case STUCLI_STUTYPE_MPACC:i=4
		Case Else
			getClientStatus=0
			Exit Function
		End Select
		getClientStatus=m_client_status(i,Opr)
	End Function
	Public Function setClientStatus(val)
		Dim arr,i,j
		arr=Split(val,",")
		For i=1 To STUTYPE_COUNT
			For j=1 To OPRTYPE_COUNT
				m_client_status(i,j)=arr(OPRTYPE_COUNT*i+j-OPRTYPE_COUNT)
			Next
		Next
	End Function
	Public Function isOpenFor(StuType,Opr)
		Dim i
		Select Case StuType
		Case STUCLI_STUTYPE_ME:i=1
		Case STUCLI_STUTYPE_MBA:i=2
		Case STUCLI_STUTYPE_EMBA:i=3
		Case STUCLI_STUTYPE_MPACC:i=4
		Case Else
			isOpenFor=False
			Exit Function
		End Select
		isOpenFor=m_client_status(i,Opr)="1" And DateDiff("d",opentime_start(Opr),Now)>=0 And DateDiff("d",opentime_end(Opr),Now)<=0
	End Function
	Public Property Get SystemStatus()
		SystemStatus=m_systemstatus
	End Property
	Public Property Let SystemStatus(val)
		m_systemstatus=val
	End Property
End Class

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

// 注册系统选项
Dim dictSystemOptions:Set dictSystemOptions = Server.CreateObject("Scripting.Dictionary")
'dictSystemOptions.Add "send_accept_email",			Array("导师确认填报时，向学生发送通知邮件", 0)

// 获取系统选项值（没有则使用默认值）
Function getSystemOption(name,stu_type_id)
	If Not dictSystemOptions.Exists(name) Then Exit Function
	Dim conn,rs,sql,result
	Connect conn
	sql="EXEC spGetSystemOption ?,?,?"
	Set rs=ExecQuery(conn,sql,Array(CmdParam("Name",adVarChar,adParamInput,50,name),_
		CmdParam("StuTypeId",adInteger,adParamInput,4,stu_type_id),_
		CmdParam("DefaultValue",adInteger,adParamInput,4,dictSystemOptions.Item(name)(1))),result)
	getSystemOption=rs(0).Value
	CloseRs rs
	CloseConn conn
End Function

// 注册脚本别名
Dim dictScripts:Set dictScripts = Server.CreateObject("Scripting.Dictionary")
dictScripts.Add "admin",	"admin"
dictScripts.Add "common",	Array("utils", "query")
dictScripts.Add "drafting",	"drafting"
dictScripts.Add "expert",	"expert"
dictScripts.Add "expertList",	"expertList"
dictScripts.Add "jeasyui",	Array("jeasyui/jquery.easyui.min", "jeasyui/easyloader")
dictScripts.Add "jquery",	"jquery-1.12.4.min"
dictScripts.Add "keywordList",	"keywordList"
dictScripts.Add "notifyList",	"notifyList"
dictScripts.Add "review",	"review"
dictScripts.Add "reviewSettings",	"reviewSettings"
dictScripts.Add "student",	"student"
dictScripts.Add "systemSettings",	"systemSettings"
dictScripts.Add "thesis",	"thesis"
dictScripts.Add "tutor",	"tutor"
dictScripts.Add "upload",	"upload"
dictScripts.Add "uploadTable",	"uploadTable"
dictScripts.Add "uploadThesis",	"uploadThesis"

// 注册样式表别名
Dim dictStylesheets:Set dictStylesheets = Server.CreateObject("Scripting.Dictionary")
dictStylesheets.Add "admin",	"admin"
dictStylesheets.Add "global",	"global"
dictStylesheets.Add "jeasyui",	Array("jeasyui/themes/icon", "jeasyui/themes/bootstrap/easyui")
dictStylesheets.Add "student",	"student"
dictStylesheets.Add "tutor",	"tutor"

Function useScript(arr)
	If Not IsArray(arr) Then arr=Array(arr)
	Dim i
	For i=0 To UBound(arr)
		Dim alias:alias=arr(i)
		If Not dictScripts.Exists(alias) Then
%><script type="text/javascript">console.warn("useScript(): 脚本别名'<%=alias%>'未注册。");</script><%
		Else
			Dim items:items=dictScripts.Item(alias)
			If Not IsArray(items) Then
%><script type="text/javascript" src="<%=baseUrl()%>scripts/<%=items%>.js"></script><%
			Else
				Dim j
				For j=0 To UBound(items)
%><script type="text/javascript" src="<%=baseUrl()%>scripts/<%=items(j)%>.js"></script><%
				Next
			End If
		End If
	Next
End Function

Function useStylesheet(arr)
	If Not IsArray(arr) Then arr=Array(arr)
	Dim i
	For i=0 To UBound(arr)
		Dim alias:alias=arr(i)
		If Not dictStylesheets.Exists(alias) Then
%><script type="text/javascript">console.warn("useStylesheet(): 样式表别名'<%=alias%>'未注册。");</script><%
		Else
			Dim items:items=dictStylesheets.Item(alias)
			If Not IsArray(items) Then
%><link href="<%=baseUrl()%>css/<%=items%>.css" rel="stylesheet" type="text/css" /><%
			Else
				Dim j
				For j=0 To UBound(items)
%><link href="<%=baseUrl()%>css/<%=items(j)%>.css" rel="stylesheet" type="text/css" /><%
				Next
			End If
		End If
	Next
End Function

Function CreateDictionary()
	Dim ret: Set ret = Server.CreateObject("Scripting.Dictionary")
	Set CreateDictionary = ret
	Set ret = Nothing
End Function

Const STUCLI_OPR_TABLE1 = 1
Const STUCLI_OPR_TABLE2 = 2
Const STUCLI_OPR_TABLE3 = 3
Const STUCLI_OPR_TABLE4 = 4
Const STUCLI_OPR_DETECT_REVIEW = 5
Const STUCLI_OPR_REVIEW = 6
Const STUCLI_OPR_MODIFY = 7
Const STUCLI_OPR_FINAL = 8
Const STUCLI_OPENTIME_START = 1
Const STUCLI_OPENTIME_END = 2
Const STUCLI_STUTYPE_ME = 5
Const STUCLI_STUTYPE_MBA = 6
Const STUCLI_STUTYPE_EMBA = 7
Const STUCLI_STUTYPE_MPACC = 9
Const STUCLI_STATUS_CLOSED = 0
Const STUCLI_STATUS_OPEN = 1
Const STUCLI_STATUS_DEBUG = 2

Dim arrStuTypeId:arrStuTypeId=Array(0,5,6,7,9)
Dim arrStuTypeName:arrStuTypeName=Array("","ME","MBA","EMBA","MPAcc")
Dim arrDiplomaName:arrDiplomaName=Array("","专科","本科","硕士研究生","博士研究生 ","博士后")

' 禁止缓存，仅在需要客户端重新请求 JS/CSS/HTML 等静态文件时使用
Response.CacheControl="no-cache"
Response.Charset="utf-8"
%>
<!--#include file="../inc/Conversion.inc" -->
<!--#include file="../inc/JSONWriter.inc" -->
<!--#include file="../inc/PinyinQuery.inc"-->