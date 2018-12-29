<!--#include file="../inc/global.inc"-->
<%
Function join(arr,delim)
	If Not IsArray(arr) Then
		join=arr
		Exit Function
	End If
	Dim i,ret
	For i=0 To UBound(arr)
		If i>0 Then ret=ret&delim
		ret=ret&arr(i)
	Next
	join=ret
End Function

Function isMatched(pattern,s)
	' 判断指定字符串是否满足指定模式
	Dim regEx:Set regEx=New RegExp
	regEx.Pattern=pattern
	isMatched=regEx.Test(s)
	Set regEx=Nothing
End Function

Function generatePassword()
	' 生成随机密码序列
	Dim i,j,lenPwdChar,lenPwd,ret
	Const pwdchar="Aa0Bb1Cc2Dd3Ee4Ff5Gg6Hh7Ii8Jj9Kk9L8Mm7Nn6Oo5Pp4Qq3Rr2Ss1Tt0UuVvWwXxYyZzAa0Bb1Cc"
	lenPwdChar=Len(pwdchar)
	lenPwd=7
	Randomize
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

Function getReviewResult(n)
	Dim ret
	Select Case n
	Case 1:ret="同意答辩"
	Case 2:ret="需做适当修改"
	Case 3:ret="需做重大修改"
	Case 4:ret="不同意答辩"
	Case 5:ret="尚未返回"
	End Select
	getReviewResult=ret
End Function

Function getFinalResult(n)
	Dim ret
	Select Case n
	Case 1:ret="同意答辩"
	Case 2:ret="适当修改"
	Case 3:ret="重大修改"
	Case 4:ret="加送两份"
	Case 5:ret="延期送审"
	Case 6:ret="暂无"
	End Select
	getFinalResult=ret
End Function

Function updateActiveTime(teacherID)
	' 更新数据库中用户使用评阅系统时间的记录
	Dim conn,sql
	Connect conn
	sql="UPDATE TEST_THESIS_REVIEW_NOTIFY_INFO SET LAST_ACTIVE_TIME="&toSqlString(Now)&" WHERE USER_ID="&teacherID
	conn.Execute sql
	CloseConn conn
	updateActiveTime=1
End Function

Function diplomaList(ctlname,sel)	' 显示学历选择框
	Dim i
%><div class="divcontrol"><select name="<%=ctlname%>"><option value="0">请选择</option><%
	For i=1 To UBound(arrDiplomaName)
%><option value="<%=i%>"<% If sel=i Then Response.Write " selected"%>><%=arrDiplomaName(i)%></option><%
	Next %>
</select></div><%
End Function

Function correlationTypeRadios(ctlname,sel)	' 显示相关程度单选按钮组
	Dim arr,i
	arr=Array("","相关","不相关")
	For i=1 To UBound(arr)
		If i>1 Then Response.Write "&emsp;"
%><label for="<%=ctlname&i%>"><input type="radio" name="<%=ctlname%>" id="<%=ctlname&i%>" value="<%=i%>"<% If sel=i Then %> checked="true"<% End If %>><%=arr(i)%></label><%
	Next
End Function

Function masterLevelRadios(ctlname,sel)	' 显示对论文内容熟悉程度单选按钮组
	Dim arr,i
	arr=Array("","很熟悉","熟悉","一般")
	For i=1 To UBound(arr)
		If i>1 Then Response.Write "&emsp;"
%><label for="<%=ctlname&i%>"><input type="radio" name="<%=ctlname%>" id="<%=ctlname&i%>" value="<%=i%>"<% If sel=i Then %> checked="true"<% End If %>><%=arr(i)%></label><%
	Next
End Function

Function reviewLevelRadios(ctlname,rev_type,sel)	' 显示对学位论文的总体评价单选按钮组
	Dim arr,i
	If rev_type=1 Then
		arr=Array("","优","良","中","差")
	Else
		arr=Array("","优秀","良好","合格","不合格")
	End If
	For i=1 To UBound(arr)
		If i>1 Then Response.Write "&emsp;"
%><label for="<%=ctlname&i%>"><input type="radio" name="<%=ctlname%>" id="<%=ctlname&i%>" value="<%=i%>"<% If sel=i Then %> checked="true"<% End If %>><%=arr(i)%></label><%
	Next
End Function

Function reviewResultRadios(ctlname,sel)	' 显示评审结果单选按钮组
	Dim arr,i
	arr=Array("","同意答辩","适当修改后答辩","需做重大修改后方可答辩","不同意答辩")
	For i=1 To UBound(arr)
		If i>1 Then Response.Write "&emsp;"
%><label for="<%=ctlname&i%>"><input type="radio" name="<%=ctlname%>" id="<%=ctlname&i%>" value="<%=i%>"<% If sel=i Then %> checked="true"<% End If %>><%=arr(i)%></label><%
	Next
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

Function finalResultList(ctlname,sel,showtip)	' 显示处理意见选择框
	Dim arr,i
	arr=Array("","I","II","III","IV","V","VI")
%><div class="divcontrol" onmousedown="return false" onkeydown="return false"><select name="<%=ctlname%>"><%
	If showtip Then %><option value="0">暂无</option><% End If
	For i=1 To UBound(arr)
%><option value="<%=i%>"<% If sel=i Then Response.Write " selected"%>><%=arr(i)%></option><%
	Next %>
</select></div><%
End Function

Function loadReviewScoringInfo(review_type,formcode,power1code,power2code)
	' 显示评阅书评价指标信息
	Dim stream,filepath
	Dim data,numScores1,numScores2,nameScore1,nameScore2,power1,power2,power2sum,detail
	Dim c,i,j,k,tmp,tmp2,tmp3
	Dim retcode,retpower1,retpower2
	Set stream=Server.CreateObject("ADODB.Stream")
	filepath=Server.MapPath("/ThesisReview/expert/data/scoreform"&review_type&".html")
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
	sql="SELECT EXPERT_NAME,PRO_DUTY_NAME,LAST_DIPLOMA,EXPERTISE,WORKPLACE,ADDRESS,MAILCODE,TELEPHONE,MOBILE,EMAIL,BANK_ACCOUNT,BANK_NAME,IDCARD_NO FROM TEST_THESIS_REVIEW_EXPERT_INFO WHERE TEACHER_ID="&Session("Tid")
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

Function getSystemStatus()
	Dim conn,rs,sql,result
	Dim sem_info:sem_info=getCurrentSemester()
	Connect conn
	sql="SELECT EXP_STARTDATE,EXP_ENDDATE FROM TEST_THESIS_REVIEW_SYSTEM WHERE USE_YEAR="&sem_info(0)&" AND USE_SEMESTER="&sem_info(1)&" AND VALID=1"
	GetRecordSetNoLock conn,rs,sql,result
	If rs.EOF Then
		getSystemStatus=0
	Else
		startdate=rs(0).Value
		enddate=rs(1).Value
		If DateDiff("d",startdate,Now)<0 Or DateDiff("d",enddate,Now)>0 Then
			getSystemStatus=1
		Else
			getSystemStatus=2
		End If
	End If
	CloseRs rs
	CloseConn conn
End Function

Dim nSystemStatus,startdate,enddate
nSystemStatus=getSystemStatus()
If nSystemStatus=0 Then
%><html><head><link href="../css/tutor.css" rel="stylesheet" type="text/css" /></head>
<body class="exp"><center><div class="content"><font color=red size="4">电子评阅系统未启用！</font><br /><input type="button" value="返 回" onclick="history.go(-1)" /></center></div></body></html><%
	Response.End()
End If
%>