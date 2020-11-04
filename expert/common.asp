<%
Function getProDutyNameOf(tid)
	Dim conn,rs,sql,num
	Connect conn
	sql="SELECT PRO_DUTYNAME FROM ViewTeacherInfo WHERE TEACHERID="&tid
	GetRecordSetNoLock conn,rs,sql,num
	If Not rs.EOF Then
		getProDutyNameOf=rs(0)
	End If
	CloseRs rs
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

Function reviewResultList(ctlname,sel,showtip)	' 显示评审结果选择框
	Dim arr,i
	arr=Array("","A","B","C","D","E")
%><select name="<%=ctlname%>"><%
	If showtip Then %><option value="0">请选择</option><% End If
	For i=1 To UBound(arr)
%><option value="<%=i%>"<% If sel=i Then Response.Write " selected"%>><%=arr(i)%></option><%
	Next %>
</select><%
End Function

Function checkIfProfileFilledIn()
	Dim conn,rs,sql,i,ret
	Connect conn
	sql="SELECT EXPERT_NAME,PRO_DUTY_NAME,LAST_DIPLOMA,EXPERTISE,WORKPLACE,ADDRESS,MAILCODE,TELEPHONE,MOBILE,EMAIL,BANK_ACCOUNT,BANK_NAME,IDCARD_NO FROM Experts WHERE TEACHER_ID="&Session("TId")
	Set rs=conn.Execute(sql)
	If rs.EOF Then
		checkIfProfileFilledIn=True
		Exit Function
	End If
	ret=True
	For i=0 To rs.Fields.Count-1
		If Len(rs(i))=0 Or IsNull(rs(i)) Then
			ret=False
			Exit For
		End If
	Next
	CloseRs rs
	CloseConn conn
	checkIfProfileFilledIn=ret
End Function

Function getSectionAccessibilityInfo(activity_id, stu_type_id, section_id)
	Dim section, time_flag, tip
	Dim accessible:accessible=False
	If Not isActivityOpen(activity_id) Then
		Set section=getSectionInfo(Null, Null, section_id)
		time_flag=-3
	Else
		Set section=getSectionInfo(activity_id, stu_type_id, section_id)
		If section Is Nothing Then
			Set section=getSectionInfo(Null, Null, section_id)
			time_flag=-2
		Else
			time_flag=compareNowWithSectionTime(section)
		End If
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
	Else
		tip=Format("【{0}】环节开放时间为{1}至{2}，当前在开放时间内，可以正常评阅。",_
			section("Name"), toDateTime(section("StartTime"), 0), toDateTime(section("EndTime"), 0))
	End If
	Dim dict:Set dict=CreateDictionary()
	dict.Add "section", section
	dict.Add "time_flag", time_flag
	dict.Add "accessible", accessible
	dict.Add "tip", tip
	Set getSectionAccessibilityInfo=dict
End Function

Dim arrFileListName,arrFileListNamePostfix,arrFileListPath,arrFileListField
arrFileListName=Array("","送审论文","论文评阅书 1","论文评阅书 2")
arrFileListNamePostfix=Array("","","论文评阅书(1)","论文评阅书(2)")
arrFileListPath=Array("","student/upload","expert/export","expert/export")
arrFileListField=Array("","THESIS_FILE2","ReviewFile1","ReviewFile2")
%>