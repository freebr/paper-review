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

Function loadReviewScoringInfo(review_type,formcode,power1code,power2code)
	' 显示评阅书评价指标信息
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
	sql="SELECT EXPERT_NAME,PRO_DUTY_NAME,EXPERTISE,WORKPLACE,ADDRESS,MAILCODE,TELEPHONE,MOBILE,EMAIL,BANK_ACCOUNT,BANK_NAME,IDCARD_NO FROM Experts WHERE TEACHER_ID="&Session("TId")
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
%>