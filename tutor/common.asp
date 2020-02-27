<%
Function generatePassword()
	' 生成随机密码序列
	Dim i,j,lenPwdChar,lenPwd,ret
	Const pwdchar="Aa0Bb1Cc2Dd3Ee4Ff5Gg6Hh7Ii8Jj9Kk9L8Mm7Nn6Oo5Pp4Qq3Rr2Ss1Tt0UuVvWwXxYyZzAa0Bb1Cc"
	lenPwdChar=Len(pwdchar)
	lenPwd=7
	Randomize()
	For i=1 To lenPwd
		j=Int(Rnd()*lenPwdChar)
		ret=ret&Mid(pwdchar,j,1)
	Next
	generatePassword=ret
End Function

Function getTeacherIdByName(name)
	If IsNull(name) Then
		getTeacherIdByName=-1
		Exit Function
	End If
	Dim conn,rsTeacher,sql,count
	name=Replace(name," ",vbNullString)
	name=Replace(name,"　",vbNullString)
	name=Replace(name,"'","''")
	name=Replace(name,"""","""""")
	Connect conn
	sql="SELECT TEACHERID,TEACHERNAME FROM TEACHER_INFO WHERE TEACHERNAME='"&name&"' AND VALID=0"
	GetRecordSetNoLock conn,rsTeacher,sql,count
	If rsTeacher.EOF Then
		getTeacherIdByName=-1
	Else
		getTeacherIdByName=rsTeacher("TEACHERID")
	End If
	CloseRs rsTeacher
	CloseConn conn
End Function

Function getProDutyNameOf(tid)
	Dim conn,rs,sql,count
	Connect conn
	sql="SELECT PRO_DUTYNAME FROM ViewTeacherInfo WHERE TEACHERID="&tid
	GetRecordSetNoLock conn,rs,sql,count
	If Not rs.EOF Then
		getProDutyNameOf=rs(0)
	End If
	CloseRs rs
	CloseConn conn
End Function

Function getNoticeText(stuType,noticeName)
	Dim conn,rs,sql,count
	Connect conn
	sql="EXEC spGetNoticeText ?,?"
	Dim ret:Set ret=ExecQuery(conn,sql,_
		CmdParam("StudentType",adInteger,4,stuType),_
		CmdParam("NoticeName",adVarWChar,50,noticeName))
	Set rs=ret("rs")
    count=ret("count")
	If rs.EOF Then
		getNoticeText="【无】"
	Else
		getNoticeText=rs(0).Value
	End If
	CloseRs rs
	CloseConn conn
End Function

Function addAuditRecord(dissertation_id,filename,audit_type,audit_time,is_passed,eval_text)
	If Len(eval_text)=0 Then
		If is_passed Then eval_text="审核通过" Else eval_text="审核不通过"
	End If
	Dim conn,sql
	Connect conn
	sql="EXEC spAddAuditRecord ?,?,?,?,?,?,?,NULL"
	ExecNonQuery conn,sql,_
		CmdParam("dissertation_id",adInteger,4,dissertation_id),_
		CmdParam("audit_file",adVarWChar,50,filename),_
		CmdParam("audit_type",adInteger,4,audit_type),_
		CmdParam("audit_time",adDate,4,audit_time),_
		CmdParam("auditor_id",adInteger,4,Session("TId")),_
		CmdParam("is_passed",adVarWChar,500,is_passed),_
		CmdParam("comment",adLongVarWChar,5000,eval_text)
	CloseConn conn
End Function

Function updateActiveTime(teacherID)
	' 更新数据库中用户使用评阅系统时间的记录
	Dim conn,sql
	Connect conn
	sql="UPDATE NotifyList SET LAST_ACTIVE_TIME="&toSqlString(Now)&" WHERE USER_ID="&teacherID
	conn.Execute sql
	CloseConn conn
	updateActiveTime=1
End Function

Function sendEmailToStudent(dissertation_id,file_type_name,is_pass,ByVal eval_text)
	If Len(eval_text)=0 Then eval_text="无"
	Dim conn:Connect conn
	Dim sql:sql="SELECT ActivityId,TEACHTYPE_ID,STU_NAME,STU_NO,CLASS_NAME,SPECIALITY_NAME,EMAIL,THESIS_SUBJECT,TUTOR_NAME,TUTOR_EMAIL FROM ViewDissertations WHERE ID=?"
	Dim ret:Set ret=ExecQuery(conn,sql,CmdParam("ID",adInteger,4,dissertation_id))
	Dim rs:Set rs=ret("rs")
	Dim activity_id:activity_id=rs("ActivityId")
	Dim stu_type:stu_type=rs("TEACHTYPE_ID")
	Dim dict:Set dict=CreateDictionary()
	Dim template_name,operation_name,is_sent
	dict("stuname")=rs("STU_NAME")
	dict("stuno")=rs("STU_NO")
	dict("stuclass")=rs("CLASS_NAME")
	dict("stuspec")=rs("SPECIALITY_NAME")
	dict("stumail")=rs("EMAIL")
	dict("subject")=rs("THESIS_SUBJECT")
	dict("tutorname")=rs("TUTOR_NAME")
	dict("tutormail")=rs("TUTOR_EMAIL")
	CloseRs rs
	CloseConn conn

	If Len(file_type_name)=0 Then
		template_name="pyyjqrtzyj"
		operation_name="确认评阅书"
	Else
		If is_pass Then
			template_name="lwshtgtzyj"
			operation_name=Format("审核通过[{0}]",file_type_name)
		Else
			template_name="lwshwtgtzyj"
			operation_name=Format("审核不通过[{0}]",file_type_name)
		End If
	End If
	is_sent=sendNotifyMail(activity_id,stu_type,template_name,dict("stumail"),dict)
	writeNotificationEventLog usertypeTutor,Session("Tname"),operation_name,usertypeStudent,_
		dict("stuname"),dict("stumail"),notifytypeMail,is_sent
	sendEmailToStudent=is_sent
End Function

Function getSectionAccessibilityInfo(activity_id, stu_type_id, section_id)
	Dim section, time_flag, tip
	Dim accessible:accessible=False
	If Not isActivityOpen(activity_id) Then
		Set section=getSectionInfo(Null, Null, section_id)
		time_flag=-3
	Else
		Set section=getSectionInfo(activity_id, stu_type_id, section_id)
		time_flag=compareNowWithSectionTime(section)
		accessible=time_flag=0
	End If
	If Not accessible Then
		If time_flag=-3 Then
			tip=Format("当前评阅活动已关闭，不能执行【{0}】操作。", section("Name"))
		ElseIf time_flag=-2 Then
			tip=Format("【{0}】环节已关闭，不能执行操作。", section("Name"))
		Else
			tip=Format("【{0}】环节开放时间为{1}至{2}，当前不在开放时间内，不能执行操作。",_
				section("Name"), toDateTime(section("StartTime"), 0), toDateTime(section("EndTime"), 0))
		End If
	End If
	Dim dict:Set dict=CreateDictionary()
	dict.Add "section", section
	dict.Add "time_flag", time_flag
	dict.Add "accessible", accessible
	dict.Add "tip", tip
	Set getSectionAccessibilityInfo=dict
End Function

If Not hasPrivilege(Session("Twriteprivileges"),"I11") And Not hasPrivilege(Session("Treadprivileges"),"I11") Then
	showErrorPage "您没有访问本系统的权限！", "提示"
End If

Dim arrTable:arrTable=Array("","开题报告表","中期检查表","预答辩申请表","答辩审批材料")
%>