<%Response.Charset="utf-8"%>
<!--#include virtual="/pub/mail.asp"-->
<!--#include file="../inc/db.asp"-->
<!--#include file="../inc/setEditor.asp"-->
<!--#include file="../inc/ckeditor/ckeditor.asp"-->
<!--#include file="../inc/ckfinder/ckfinder.asp"-->
<!--#include file="../inc/global.inc"-->
<%Response.Expires=-1
If IsEmpty(Session("user")) Then Response.Redirect("../error.asp?timeout")
curstep=Request.QueryString("step")
If Len(curstep)=0 Or Not IsNumeric(curstep) Then curstep="1"
sem_info=getCurrentSemester()
Dim arrMailId,arrMailSubject,arrStuOprFlag,arrStuOprName,arrStuType
Dim tutor_startdate,tutor_enddate,exp_startdate,exp_enddate
Dim stu_startdate():ReDim stu_startdate(OPRTYPE_COUNT)
Dim stu_enddate():ReDim stu_enddate(OPRTYPE_COUNT)
Dim stu_clientstatus
arrStuOprFlag=Array("","TABLE1","TABLE2","TABLE3","TABLE4","DETECT","REVIEW","MODIFY","FINAL")
arrStuOprName=Array("","开题报告表","中期检查表","预答辩申请表","答辩审批材料","送检论文","送审论文","答辩论文","定稿论文")
arrStuType=Array("","ME","MBA","EMBA","MPAcc")
arrMailSubject=Array("","论文送审通知邮件（学生）","论文送审通知邮件（导师）","论文待评阅通知邮件","论文待评阅通知短信","论文审核通知邮件","论文审核未通过通知邮件","论文审核通过通知邮件","评阅意见确认通知邮件","信息导入通知邮件（学生）","信息导入通知邮件（导师）","待办事项通知邮件")
ReDim arrMailId(UBound(arrMailSubject))
If curstep="1" Then
	ReDim stu_clientstatus(OPRTYPE_COUNT*STUTYPE_COUNT)
	Connect conn
	sql="SELECT * FROM TEST_THESIS_REVIEW_SYSTEM WHERE USE_YEAR="&sem_info(0)&" AND USE_SEMESTER="&sem_info(1)
	GetRecordSetNoLock conn,rs,sql,result
	If result=0 Then	' 本学期无系统设置
		For i=1 To OPRTYPE_COUNT
			stu_startdate(i)=Date
			stu_enddate(i)=Date+7
		Next
		exp_startdate=Date+7
		exp_enddate=Date+14
		tutor_startdate=Date
		tutor_enddate=Date+7
		turn_num=1
		isValid=False
	Else	' 本学期有系统设置
		bSet=True
		For i=1 To OPRTYPE_COUNT
			stu_startdate(i)=rs("STU_"&arrStuOprFlag(i)&"_STARTDATE")
			stu_enddate(i)=rs("STU_"&arrStuOprFlag(i)&"_ENDDATE")
		Next
		exp_startdate=rs("EXP_STARTDATE")
		exp_enddate=rs("EXP_ENDDATE")
		tutor_startdate=rs("TUTOR_STARTDATE")
		tutor_enddate=rs("TUTOR_ENDDATE")
		stu_clientstatus=Split(rs("STU_CLIENT_STATUS"),",")
		turn_num=rs("TURN_NUM")
		For i=1 To UBound(arrMailSubject)
			arrMailId(i)=rs("MAIL_"&i)
		Next
		isValid=rs("VALID")
	End If
	CloseRs rs
	If Not bSet Then
		sql="SELECT TOP 1 * FROM TEST_THESIS_REVIEW_SYSTEM ORDER BY ID DESC"
		GetRecordSetNoLock conn,rs,sql,result
		If result Then ' 显示最新学期的邮件内容
			For i=1 To UBound(arrMailId)
				arrMailId(i)=rs("MAIL_"&i)
			Next
		End If
	End If
	CloseRs rs
	CloseConn conn
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="../css/admin.css" rel="stylesheet" type="text/css" />
<script src="../scripts/jquery-1.6.3.min.js" type="text/javascript"></script>
<script src="../scripts/query.js" type="text/javascript"></script>
<script src="../scripts/utils.js" type="text/javascript"></script>
<script src="../scripts/systemSettings.js" type="text/javascript"></script>
</head>
<body bgcolor="ghostwhite">
<center><font size=4><b><%=sem_info(0)%>-<%=sem_info(0)+1%>年度<%=sem_info(2)%>学期专业学位论文评阅系统设置</b><br><%
	If Not bSet Then
%><span style="color:red;font-weight:bold">(请先设置开放时间等属性，方可开放系统)</span><%
	End If %></font>
<form id="fmSettings" id="fmSettings" method="post" action="?step=2" onsubmit="return chkForm()">
<input type="hidden" name="In_PERIOD_ID" value="<%=sem_info(3)%>">
<table width="900" cellpadding="2" cellspacing="1" bgcolor="dimgray">
<tr bgcolor="ghostwhite">
<td align="left"><%
	Select Case Request("ok")
	Case "0"
%><span style="color:red">更新设置时发生错误，请检查设置参数是否正确</span><br/><%
	Case "1"
%><span style="color:red">设置成功</span><br/><%
	End Select
%></td></tr>
<tr bgcolor="ghostwhite">
<td align="center">当前系统状态：<%
	If isValid Then
%>开放<br/><input type="button" name="btnOpenSystem" value="关闭评阅系统" onclick="this.disabled=true;location.href='setSystemStatus.asp?open=1'" /><%
	Else
%>关闭<br/><input type="button" name="btnOpenSystem" value="开放评阅系统" <% If Not bSet Then %>disabled <% Else %>onclick="this.disabled=true;location.href='setSystemStatus.asp?open=0'" <% End If %>/><%
	End If %><input type="button" value="del" onclick="location.href='setSystemStatus.asp?open=101'" style="display:none" /></td></tr>
</td></tr>
<tr bgcolor="ghostwhite">
<td align="center"><p>当前是第&nbsp;<input type="text" name="turn_num" id="turn_num" value="<%=turn_num%>" size="5" style="text-align:center" />&nbsp;批
&nbsp;<input type="button" id="btnnewturn" value="增加批次" /><br />
<input type="button" id="btnexport" value="导出本批次评审结果" />
&nbsp;<input type="button" name="viewexportfiles" value="查看以往批次评审结果" onclick="window.open('/ThesisReview/admin/export/spec')" /><br />
&emsp;&emsp;&emsp;专家端开放时间：<input type="text" name="exp_startdate" id="exp_startdate" class="date" value="<%=exp_startdate%>" title="专家端开放时间起始日期" />
<img style="cursor:pointer" src="../images/calendar.gif" onclick="showCalendar('exp_startdate','<%=exp_startdate%>')" title="打开日历">&nbsp;至&nbsp;
<input type="text" name="exp_enddate" id="exp_enddate" class="date" value="<%=exp_enddate%>" title="专家端开放时间截止日期" />
<img style="cursor:pointer" src="../images/calendar.gif" onclick="showCalendar('exp_enddate','<%=exp_enddate%>')" title="打开日历"><br />
&emsp;&emsp;&emsp;导师端开放时间：<input type="text" name="tutor_startdate" id="tutor_startdate" class="date" value="<%=tutor_startdate%>" title="导师端开放时间起始日期" />
<img style="cursor:pointer" src="../images/calendar.gif" onclick="showCalendar('tutor_startdate','<%=tutor_startdate%>')" title="打开日历">&nbsp;至&nbsp;
<input type="text" name="tutor_enddate" id="tutor_enddate" class="date" value="<%=tutor_enddate%>" title="导师端开放时间截止日期" />
<img style="cursor:pointer" src="../images/calendar.gif" onclick="showCalendar('tutor_enddate','<%=tutor_enddate%>')" title="打开日历"></p>
<p><table width="800" cellpadding="0" cellspacing="0" border="0">
<tr bgcolor="ghostwhite"><td align="center" colspan="3">学生端上传通道开放时间和开放对象：</td></tr>
<%
	For i=1 To OPRTYPE_COUNT
%><tr bgcolor="ghostwhite">
<td align="right"><%=arrStuOprName(i)%>：</td>
<td align="center"><input type="text" name="stu_startdate<%=i%>" id="stu_startdate<%=i%>" class="date" value="<%=stu_startdate(i)%>" title="学生端<%=arrStuOprName(i)%>环节开放时间起始日期" />
<img style="cursor:pointer" src="../images/calendar.gif" onclick="showCalendar('stu_startdate<%=i%>','<%=stu_startdate(i)%>')" title="打开日历">&nbsp;至&nbsp;
<input type="text" name="stu_enddate<%=i%>" id="stu_enddate<%=i%>" class="date" value="<%=stu_enddate(i)%>" title="学生端<%=arrStuOprName(i)%>环节开放时间截止日期" />
<img style="cursor:pointer" src="../images/calendar.gif" onclick="showCalendar('stu_enddate<%=i%>','<%=stu_enddate(i)%>')" title="打开日历"></td>
<td align="center"><%
		For j=1 To STUTYPE_COUNT
			k=OPRTYPE_COUNT*j+i-OPRTYPE_COUNT
			If stu_clientstatus(k)="1" Then
				checkflag="checked=""true"" "
			Else
				checkflag=vbNullString
			End If
%>&emsp;<label for="stu_clientstatus<%=k%>"><input type="checkbox" name="stu_clientstatus<%=k%>" id="stu_clientstatus<%=k%>" <%=checkflag%>/><%=arrStuType(j)%></label><%
		Next
%></td></tr><%
	Next %></table></p></td></tr>
<tr bgcolor="ghostwhite"><td colspan=4><p align="center"><input type="submit" value="更改设置"></p></td></tr>
<tr bgcolor="ghostwhite"><td align="left"><font style="font-weight:bold">通知邮件/短信内容设置</font></td></tr>
<tr bgcolor="ghostwhite"><td align="left"><select id="maillist" onchange="switchMailContent(this.selectedIndex)"><option>【请选择】</option><%
	For i=1 To UBound(arrMailSubject)
%><option><%=arrMailSubject(i)%></option><%
	Next
%></select><br />
<span id="mailtip">字段符号:<br/>$stuname - 学生姓名,$stuno - 学号,$stuclass - 学生班级,$stuspec - 所选专业,$stumail - 学生邮箱,<br/>
$subject - 论文题目,$tutorname - 导师姓名,$tutormail - 导师邮箱,$expertname - 专家姓名,<br/>$filename - 审核文件名称/意见类型,$uploadtime - 审核文件上传时间,$evaltext - 导师意见,$postscript - 附注</span></td></tr>
<tr bgcolor="ghostwhite">
<td align="left"><%
	For i=1 To UBound(arrMailSubject)
%><div id="divmailcontent<%=i%>" style="display:none"><% SetEditorWithName "mailcontent"&i,getEmailTemplateContent(arrMailId(i)),170 %></div><%
	Next %>
</td></tr>
<tr bgcolor="ghostwhite"><td align="center"><input type="submit" value="更改设置"></td></tr>
</table></form>
</center></body>
<script type="text/javascript">
	$('#btnnewturn').click(function() {
		$(this).val('正在执行，请稍候……').attr('disabled',true);
		$(':input[name="turn_num"]').val('<%=turn_num+1%>');
		this.form.submit();
	}).attr('disabled',false);
	$('#btnexport').click(function() {
		$(this).val('正在导出，请稍候……').attr('disabled',true);
		this.form.action='exportReviewStats.asp?fn=<%=sem_info(3)%>_<%=turn_num%>&turn=<%=turn_num%>';
		this.form.submit();
	}).attr('disabled',false);
</script></html><%
Else
	Dim mail_content,fieldlist
	Dim strTmp
	ReDim mail_content(UBound(arrMailSubject))
	For i=1 To OPRTYPE_COUNT
		stu_startdate(i)=Request.Form("stu_startdate"&i)
		stu_enddate(i)=Request.Form("stu_enddate"&i)
		If Len(stu_startdate(i))=0 Then stu_startdate(i)=Null
		If Len(stu_enddate(i))=0 Then stu_enddate(i)=Null
	Next
	exp_startdate=Request.Form("exp_startdate")
	exp_enddate=Request.Form("exp_enddate")
	If Len(exp_startdate)=0 Then exp_startdate=Null
	If Len(exp_enddate)=0 Then exp_enddate=Null
	tutor_startdate=Request.Form("tutor_startdate")
	tutor_enddate=Request.Form("tutor_enddate")
	If Len(tutor_startdate)=0 Then tutor_startdate=Null
	If Len(tutor_enddate)=0 Then tutor_enddate=Null
	stu_clientstatus="0"
	For i=1 To OPRTYPE_COUNT*STUTYPE_COUNT
		stu_clientstatus=stu_clientstatus&","
		If Request.Form("stu_clientstatus"&i)="on" Then
			stu_clientstatus=stu_clientstatus&"1"
		Else
			stu_clientstatus=stu_clientstatus&"0"
		End If
	Next
	turn_num=Request.Form("turn_num")
	If Not IsNumeric(turn_num) Then
		turn_num=1
	ElseIf turn_num<1 Then
		turn_num=1
	End If
	turn_num=Int(turn_num)
	For i=1 To UBound(arrMailSubject)
		arrMailId(i)=0
		mail_content(i)=Request.Form("mailcontent"&i)
	Next
	arrFieldList=Array("","$stuname,$stuno,$stuclass,$stuspec,$stumail,$subject,$tutorname,$tutormail",_
								"$stuname,$stuno,$stuclass,$stuspec,$stumail,$subject,$tutorname,$tutormail",_
								"$subject,$expertname,$expertmob,$expertmail,$postscript",_
								"$subject,$expertname,$expertmob,$expertmail,$postscript",_
								"$stuname,$stuno,$stuclass,$stuspec,$stumail,$subject,$tutorname,$tutormail,$filename,$uploadtime",_
								"$stuname,$stuno,$stuclass,$stuspec,$stumail,$subject,$tutorname,$tutormail,$filename,$evaltext",_
								"$stuname,$stuno,$stuclass,$stuspec,$stumail,$subject,$tutorname,$tutormail,$filename,$evaltext",_
								"$stuname,$stuno,$stuclass,$stuspec,$stumail,$subject,$tutorname,$tutormail,$evaltext",_
								"$stuname,$stuno,$stuclass,$stuspec,$stumail,$subject,$tutorname,$tutormail,$filename",_
								"$stuname,$stuno,$stuclass,$stuspec,$stumail,$subject,$tutorname,$tutormail,$filename",_
								"$tutorname,$tutormail,$postscript")
	
	Connect conn
	sql="SELECT * FROM TEST_THESIS_REVIEW_SYSTEM WHERE USE_YEAR="&sem_info(0)&" AND USE_SEMESTER="&sem_info(1)
	GetRecordSet conn,rs,sql,result
	On Error Resume Next
	If result=0 Then
		rs.AddNew()
		rs("USE_YEAR")=sem_info(0)
		rs("USE_SEMESTER")=sem_info(1)
	End If
	For i=1 To OPRTYPE_COUNT
		rs("STU_"&arrStuOprFlag(i)&"_STARTDATE")=stu_startdate(i)
		rs("STU_"&arrStuOprFlag(i)&"_ENDDATE")=stu_enddate(i)
	Next
	rs("EXP_STARTDATE")=exp_startdate
	rs("EXP_ENDDATE")=exp_enddate
	rs("TUTOR_STARTDATE")=tutor_startdate
	rs("TUTOR_ENDDATE")=tutor_enddate
	rs("STU_CLIENT_STATUS")=stu_clientstatus
	rs("TURN_NUM")=turn_num
	For i=1 To UBound(arrMailSubject)
		arrMailId(i)=rs("MAIL_"&i)
		If result=0 Or IsNull(arrMailId(i)) Then arrMailId(i)=0
	Next
	rs.Update()
	
	If Err.Number=0 Then ok=1 Else ok=0
	On Error GoTo 0
	
	If ok Then
		sql="UPDATE TEST_THESIS_REVIEW_SYSTEM SET "
		For i=1 To UBound(arrMailSubject)
			template_name=sem_info(0)&"-"&(sem_info(0)+1)&"年度"&sem_info(2)&"学期"&arrMailSubject(i)
			arrMailId(i)=updateEmailTemplate(arrMailId(i),template_name,arrMailSubject(i),mail_content(i),arrFieldList(i))
			If i>1 Then sql=sql&","
			sql=sql&"MAIL_"&i&"="&arrMailId(i)
		Next
		sql=sql&" WHERE USE_YEAR="&sem_info(0)&" AND USE_SEMESTER="&sem_info(1)
		conn.Execute sql
	End If
	
	CloseRs rs
	CloseConn conn
%><body><form id="ret" method="post" action="?step=1"><input type="hidden" name="ok" value="<%=ok%>" />
<input type="hidden" name="In_PERIOD_ID" value="<%=sem_info(3)%>">
</form>
<script type="text/javascript">
document.all.ret.submit();
</script></body><%
End If
%>