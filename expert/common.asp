<%
Function getTeacherIdByName(name)
	If IsNull(name) Then
		getTeacherIdByName=-1
		Exit Function
	End If
	Dim conn,rsTeacher,sql,num
	name=Replace(name," ",vbNullString)
	name=Replace(name,"　",vbNullString)
	name=Replace(name,"'","''")
	name=Replace(name,"""","""""")
	Connect conn
	sql="SELECT TEACHERID,TEACHERNAME FROM TEACHER_INFO WHERE TEACHERNAME='"&name&"' AND VALID=0"
	GetRecordSetNoLock conn,rsTeacher,sql,num
	If rsTeacher.EOF Then
		getTeacherIdByName=-1
	Else
		getTeacherIdByName=rsTeacher("TEACHERID")
	End If
	CloseRs rsTeacher
	CloseConn conn
End Function

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

Function loadReviewScoringInfo(review_type,formcode,power1code,power2code)
	' 显示评阅书评价指标信息
	Dim stream,filepath
	Dim data,numScores1,numScores2,nameScore1,nameScore2,power1,power2,power2sum,detail
	Dim c,i,j,k,tmp,tmp2,tmp3
	Dim retcode,retpower1,retpower2
	Set stream=Server.CreateObject("ADODB.Stream")
	filepath=Server.MapPath("/PaperReview/expert/data/scoreform"&review_type&".html")
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
			nameScore2=data(c)
			detail=Split(data(c+1),";")
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
	loadReviewScoringInfo=1
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
	End If
	Dim dict:Set dict=CreateDictionary()
	dict.Add "section", section
	dict.Add "time_flag", time_flag
	dict.Add "accessible", accessible
	dict.Add "tip", tip
	Set getSectionAccessibilityInfo=dict
End Function
%>