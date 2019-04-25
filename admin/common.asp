<!--#include file="../inc/global.inc"-->
<!--#include file="../inc/pinyin.asp"-->
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
	Const pwdchar="A0B1C2D3E4F5G6H7I8J9K9L8M7N6O5P4Q3R2S1T0UVWXYZA0B1C2"
	lenPwdChar=Len(pwdchar)
	lenPwd=7
	Randomize
	For i=1 To lenPwd
		j=Int(Rnd()*lenPwdChar)
		ret=ret&Mid(pwdchar,j,1)
	Next
	generatePassword=ret
End Function

Function getDateTimeId(t)
	getDateTimeId=Replace((Year(t)-2000)&FormatNumber(Month(t)/100,2)&FormatNumber(Day(t)/100,2),".","")
End Function

Function getTeachTypeIdByName(teachtype_name)
	Dim ret
	Select Case UCase(teachtype_name)
	Case "MBA":ret=6
	Case "ME":ret=5
	Case "EMBA":ret=7
	Case "MPACC":ret=9
	Case Else:ret=0
	End Select
	getTeachTypeIdByName=ret
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
	ConnectOriginDb conn
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

Function getDefenceResult(result_text)
	If IsNull(result_text) Then
		getDefenceResult=0
		Exit Function
	End If
	Dim ret
	Select Case Trim(result_text)
	Case "优秀":ret=1
	Case "良好":ret=2
	Case "及格":ret=3
	Case "不及格":ret=4
	Case Else:ret=0
	End Select
	getDefenceResult=ret
End Function

Function getReviewFileStatTxtArray()
	getReviewFileStatTxtArray=Array("不向导师和学生显示","仅向导师显示","仅向学生显示","导师和学生均显示")
End Function

Function getAdminType(userId)
	Dim conn,rs,sql,num,dict
	Connect conn
	sql="EXEC spGetEduAdminType ?"
	Set rs=ExecQuery(conn,sql,CmdParam("UserID",adInteger,adParamInput,4,userId),num)
	Set dict=Server.CreateObject("Scripting.Dictionary")
	Dim adminType,i
	If rs.EOF Then
		adminType=0
	Else
		adminType=rs("AdminType").Value
	End If
	dict.Add "", adminType
	For i=1 To UBound(arrStuTypeName)
		If (adminType And 2^i) <> 0 Then
			dict.Add arrStuTypeName(i), True
		End If
	Next
	Set getAdminType=dict
	CloseRs rs
	CloseConn conn
End Function

Function setAdminType(userId,arrTypeId)
	Dim adminType,i
	adminType=0
	For i=0 To UBound(arrTypeId)
		adminType=adminType+2^Int(arrTypeId(i))
	Next
	Dim conn,sql
	Connect conn
	sql="EXEC spSetEduAdminType ?,?"
	ExecNonQuery conn,sql,Array(CmdParam("UserID",adInteger,adParamInput,4,userId), _
		CmdParam("AdminType",adInteger,adParamInput,4,adminType))
	CloseConn conn
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

Function setNoticeText(stuType,noticeName,noticeContent)
	Dim conn,sql
	Connect conn
	sql="EXEC spSetNoticeText ?,?,?"
	ExecNonQuery conn,sql,Array(CmdParam("StudentType",adInteger,adParamInput,4,stuType), _
		CmdParam("NoticeName",adVarWChar,adParamInput,50,noticeName), _
		CmdParam("NoticeContent",adVarWChar,adParamInput,65535,noticeContent))
	CloseConn conn
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
		logtxt="行政人员["&Session("name")&"]确认评阅书，发送邮件给学生["&stuname&":"&stumail&"]"
	Else
		If ispass Then
			mailid=7
			resulttxt="通过"
		Else
			mailid=6
			resulttxt="不通过"
		End If
		fieldval=Array(stuname,stuno,stuclass,stuspec,stumail,subject,tutorname,tutormail,filetypename,evaltext)
		logtxt="行政人员["&Session("name")&"]审核"&resulttxt&"["&filetypename&"]，发送邮件给学生["&stuname&":"&stumail&"]"
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

Function semesterList(ctlname,sel)	' 显示学期选择框
	Dim conn,comm,pmSem,rs
	Dim period_id
	Connect conn
	Set comm=Server.CreateObject("ADODB.Command")
	comm.ActiveConnection=conn
	comm.CommandText="spGetSemesterList"
	comm.CommandType=adCmdStoredProc
	Set pmSem=comm.CreateParameter("semester",adInteger,adParamInput,5,0)
	comm.Parameters.Append pmSem
	Set rs=comm.Execute()
	%><select id="<%=ctlname%>" name="<%=ctlname%>"><option value="0">请选择</option><%
	Do While Not rs.EOF
		period_id=rs("PERIOD_ID").Value
%><option value="<%=period_id%>"<% If sel=period_id Then %> selected<% End If %>><%=rs("PERIOD_NAME").Value%></option><%
		rs.MoveNext()
	Loop
	Set pmSem=Nothing
	Set comm=Nothing
	CloseRs rs
	CloseConn conn
	%></select><%
End Function

Function diplomaList(ctlname,sel)	' 显示学历选择框
	Dim i
%><div class="divcontrol"><select name="<%=ctlname%>"><option value="0">请选择</option><%
	For i=1 To UBound(arrDiplomaName)
%><option value="<%=i%>"<% If sel=i Then Response.Write " selected"%>><%=arrDiplomaName(i)%></option><%
	Next %>
</select></div><%
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

Sub outputNumber(val)
	If val=0 Then Exit Sub
	Response.Write val
End Sub

Dim arrTable,reportBaseDir
arrTable=Array("","开题报告表","中期检查表","预答辩申请表","答辩审批材料")
reportBaseDir="/ThesisReview/admin/upload/report/"
%>