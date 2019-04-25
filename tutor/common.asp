<!--#include file="../inc/global.inc"-->
<%
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

Function getNoticeText(stuType,noticeName)
	Dim conn,rs,sql,num
	Connect conn
	sql="EXEC spGetNoticeText ?,?"
	Set rs=ExecQuery(conn,sql,Array(CmdParam("StudentType",adInteger,adParamInput,4,stuType),CmdParam("NoticeName",adVarWChar,adParamInput,50,noticeName)),num)
	If rs.EOF Then
		getNoticeText="【无】"
	Else
		getNoticeText=rs(0).Value
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

Function sendEmailToStudent(thesisID,filetypename,ispass,ByVal evaltext)
	Dim conn,rs,sql,num
	Dim arrMailId,mailid
	Dim stuname,stuno,stuclass,stuspec,stumail,subject,tutorname,tutormail,resulttxt,fieldval,bSuccess,logtxt
	arrMailId=getThesisReviewSystemMailIdByType(Now)
	Connect conn
	sql="SELECT * FROM ViewThesisInfo WHERE ID="&thesisID
	GetRecordSetNoLock conn,rs,sql,num
	If rs.EOF Then
		CloseRs rs
		CloseConn conn
		sendEmailToStudent=0
		Exit Function
	End If
	stuname=rs("STU_NAME")
	stuno=rs("STU_NO")
	stuclass=rs("CLASS_NAME")
	stuspec=rs("SPECIALITY_NAME")
	stumail=rs("EMAIL")
	subject=rs("THESIS_SUBJECT")
	tutorname=rs("TUTOR_NAME")
	tutormail=rs("TUTOR_EMAIL")
	If Len(evaltext)=0 Then evaltext="无"
	If Len(filetypename)=0 Then
		mailid=8
		fieldval=Array(stuname,stuno,stuclass,stuspec,stumail,subject,tutorname,tutormail,evaltext)
		logtxt="教师["&Session("Tname")&"]确认评阅书，发送邮件给学生["&stuname&":"&stumail&"]"
	Else
		If ispass Then
			mailid=7
			resulttxt="通过"
		Else
			mailid=6
			resulttxt="不通过"
		End If
		fieldval=Array(stuname,stuno,stuclass,stuspec,stumail,subject,tutorname,tutormail,filetypename,evaltext)
		logtxt="教师["&Session("Tname")&"]审核"&resulttxt&"["&filetypename&"]，发送邮件给学生["&stuname&":"&stumail&"]"
	End If
	bSuccess=sendAnnouncementEmail(arrMailId(mailid),stumail,fieldval)
	If bSuccess Then
		logtxt=logtxt&"成功。"
	Else
		logtxt=logtxt&"失败。"
	End If
	writeLog logtxt
	CloseRs rs
	CloseConn conn
	sendEmailToStudent=1
End Function

Function correlationTypeRadios(ctlname,sel)	' 显示相关程度单选按钮组
	Dim arr,i
	arr=Array("","相关","不相关")
	For i=1 To UBound(arr)
		If i>1 Then Response.Write "&emsp;"
%><input type="radio" name="<%=ctlname%>" id="<%=ctlname&i%>" value="<%=i%>"<% If sel=i Then %> checked="true"<% End If %>><label for="<%=ctlname&i%>"><%=arr(i)%></label><%
	Next
End Function

Function masterLevelRadios(ctlname,sel)	' 显示对论文内容熟悉程度单选按钮组
	Dim arr,i
	arr=Array("","很熟悉","熟悉","一般")
	For i=1 To UBound(arr)
		If i>1 Then Response.Write "&emsp;"
%><input type="radio" name="<%=ctlname%>" id="<%=ctlname&i%>" value="<%=i%>"<% If sel=i Then %> checked="true"<% End If %>><label for="<%=ctlname&i%>"><%=arr(i)%></label><%
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
%><input type="radio" name="<%=ctlname%>" id="<%=ctlname&i%>" value="<%=i%>"<% If sel=i Then %> checked="true"<% End If %>><label for="<%=ctlname&i%>"><%=arr(i)%></label><%
	Next
End Function

Function reviewResultList(ctlname,sel,showtip)	' 显示评审结果选择框
	Dim arr,i
	arr=Array("","A","B","C","D","E")
%><select name="<%=ctlname%>"><%
	If showtip Then %><option value="0">暂无</option><% End If
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

Function defenceResultList(ctlname,sel)	' 显示答辩成绩选择框
	Dim arr,i
	arr=Array("未录入","优秀","良好","及格","不及格")
%><div class="divcontrol"><select name="<%=ctlname%>"><%
	For i=0 To UBound(arr)
%><option value="<%=i%>"<% If sel=i Then Response.Write " selected"%>><%=arr(i)%></option><%
	Next %>
</select></div><%
End Function

Function grantDegreeList(ctlname,ByVal sel)	' 显示是否同意授予学位选择框
	Dim arr,i
	arr=Array("未录入","是","否")
	sel=Int(sel)+2
%><select name="<%=ctlname%>"><%
	For i=0 To UBound(arr)
%><option value="<%=i%>"<% If sel=i Then Response.Write " selected"%>><%=arr(i)%></option><%
	Next %>
</select><%
End Function

Function getSystemStatus()
	Dim conn,rs,sql,result
	Dim sem_info:sem_info=getCurrentSemester()
	Connect conn
	sql="SELECT TUTOR_STARTDATE,TUTOR_ENDDATE FROM SystemSettings WHERE USE_YEAR="&sem_info(0)&" AND USE_SEMESTER="&sem_info(1)&" AND VALID=1"
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

If Not hasPrivilege(Session("Twriteprivileges"),"I11") And Not hasPrivilege(Session("Treadprivileges"),"I11") Then
%><body bgcolor="ghostwhite"><center><font color=red size="4">您没有权限！</font><br /><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
	Response.End()
End If

Dim nSystemStatus,startdate,enddate
nSystemStatus=getSystemStatus()
If nSystemStatus=0 Then
%><body bgcolor="ghostwhite"><center><font color=red size="4">电子评阅系统未启用！</font><br /><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
	Response.End()
End If

Dim arrTable
arrTable=Array("","开题报告表","中期检查表","预答辩申请表","答辩审批材料")
%>