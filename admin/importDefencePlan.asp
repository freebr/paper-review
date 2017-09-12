<%Response.Charset="utf-8"%>
<!--#include file="../inc/upload_5xsoft.inc"-->
<!--#include file="../inc/db.asp"-->
<!--#include file="common.asp"--><%
curStep=Request.QueryString("step")
Select Case curStep
Case vbNullstring ' 文件选择页面
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="../css/admin.css" rel="stylesheet" type="text/css" />
</head>
<body bgcolor="ghostwhite">
<center><font size=4><b>导入答辩安排信息自EXCEL文件</b><br />
<form id="fmUpload" action="?step=2" method="POST" enctype="multipart/form-data">
<p><label for="chksendemail"><input type="checkbox" name="sendemail" id="chksendemail" checked />导入后发送通知邮件给导师和学生</label></p>
<p>请选择要导入的 Excel 文件：<br />文件名：<input type="file" name="excelFile" size="100" /><br />
<a href="upload/defenceplan_template.xlsx" target="_blank">点击下载答辩安排信息表格模板</a><br />
<input type="submit" name="btnsubmit" value="提 交" />&nbsp;
<input type="button" name="btnret" value="返 回" onclick="history.go(-1)" /></p></form></center>
<script type="text/javascript">
	$('#fmUpload').onsubmit=function() {
		var fileName = this.value;
		var fileExt = fileName.substring(fileName.lastIndexOf('.')).toLowerCase();
		if (fileExt != ".xls" && fileExt != ".xlsx") {
			alert("所选文件不是 Excel 文件！");
			this.form.reset();
			return false;
		}
	}
</script></body></html><%
Case 2	' 上传进程

	Dim fso,Upload,File
	
	Set Upload=New upload_5xsoft
	Set file=Upload.File("excelFile")
	send_email=Upload.Form("sendemail")
	Set fso=Server.CreateObject("Scripting.FileSystemObject")
	
	' 检查上传目录是否存在
	strUploadPath = Server.MapPath("upload\xls")
	If Not fso.FolderExists(strUploadPath) Then fso.CreateFolder(strUploadPath)
	
	fileExt=LCase(file.FileExt)
	If period_id="0" Then
		bError = True
		errstring = "请选择学期！"
	ElseIf fileExt <> "xls" And fileExt <> "xlsx" Then	' 不被允许的文件类型
		bError = True
		errstring = "所选择的不是 Excel 文件！"
	Else
		' 生成日期格式文件名
		fileid = FormatDateTime(Now(),1)&Int(Timer)
		strDestFile = fileid&"."&fileExt
		strDestPath = Server.MapPath("upload")&"\xls\"&strDestFile
		byteFileSize = file.FileSize
		' 保存
		file.SaveAs strDestPath
	End If
	Set file=Nothing
	Set Upload=Nothing
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="theme-color" content="#2D79B2" />
<title>导入答辩安排信息自EXCEL文件</title>
<link href="../css/admin.css" rel="stylesheet" type="text/css" />
</head>
<body bgcolor="ghostwhite">
<center><br /><b>导入答辩安排信息自EXCEL文件</b><br /><br /><%
	If Not bError Then %>
<form id="fmUploadFinish" action="?step=3" method="POST">
<input type="hidden" name="sendemail" value="<%=send_email%>" />
<input type="hidden" name="filename" value="<%=strDestFile%>" />
<p><%=byteFileSize%> 字节已上传，正在导入答辩安排信息...</p></form>
<script type="text/javascript">setTimeout("$('#UploadFinish').submit()",500);</script><%
	Else
%>
<script type="text/javascript">alert("<%=errstring%>");history.go(-1);</script><%
	End If
%></center></body></html><%
Case 3	' 数据读取，导入到数据库
	
	Function transchar(ByVal s)
		' 将 | 号替换为实体符号
		s=Replace(s,"|","&brvbar;")
		transchar=s
	End Function
	
	Function addData()
		' 添加数据并发送通知邮件
		Dim sql,sql2,conn,result,rsa,rsb
		Dim numInsert,numUpdate:numInsert=0:numUpdate=0
		Dim thesisid,stuid,stuid_string:numThesis=0:stuid_string="0"
		Dim member_desc,last_val(7)
		Connect conn
		Do While Not rs.EOF
			If IsNull(rs(0)) Then Exit Do
			' 按学号检索
			sql="SELECT ID,STU_ID FROM VIEW_TEST_THESIS_REVIEW_INFO WHERE VALID=1 AND STU_NO="&toSqlString(rs(2))&" AND TEACHTYPE_ID="&getTeachTypeIdByName(rs(0))&" ORDER BY STU_ID DESC"
			GetRecordSetNoLock conn,rsa,sql,result
			If rsa.EOF Then
				bError=True
				errMsg=errMsg&"学生不存在:"""&rs(1)&"""。"&vbNewLine
			Else
				numThesis=numThesis+1
				thesisid=rsa("ID")
				stuid=rsa("STU_ID")
				stuid_string=stuid_string&","&stuid
				sql="SELECT * FROM TEST_THESIS_DEFENCE_INFO WHERE THESIS_ID="&thesisid
				GetRecordSet conn,rsb,sql,result
				If rsb.EOF Then
					rsb.AddNew()
					numInsert=numInsert+1
				Else
					numUpdate=numUpdate+1
				End If
				
				member_desc=transchar(rs(4))&"|"&transchar(rs(5))&"|"&transchar(rs(6))
				rsb("THESIS_ID")=thesisid
				rsb("DEFENCE_TIME")=rs(7)
				rsb("DEFENCE_PLACE")=rs(8)
				rsb("DEFENCE_MEMBER")=member_desc
				rsb("MEMO")=rs(9)
				rsb.Update()
				CloseRs rsb
			End If
			CloseRs rsa
			rs.MoveNext()
		Loop
		
		If send_email Then
			Dim stuname,stuno,stuclass,stuspec,stumail,subject,tutorname,tutormail,filename,fieldval,bSuccess
			Dim logtxt:logtxt="行政人员["&Session("name")&"]执行答辩安排信息导入操作。"
			Dim mail_id:mail_id=getThesisReviewSystemMailIdByType(Now)
			' 批量发送通知邮件
			sql="SELECT STU_NAME,STU_NO,CLASS_NAME,SPECIALITY_NAME,EMAIL,THESIS_SUBJECT,TUTOR_NAME,TUTOR_EMAIL FROM VIEW_TEST_THESIS_REVIEW_INFO WHERE VALID=0 AND STU_ID IN ("&stuid_string&")"
			GetRecordSetNoLock conn,rsa,sql,result
			Do While Not rsa.EOF
				stuname=rsa("STU_NAME")
				stuno=rsa("STU_NO")
				stuclass=rsa("CLASS_NAME")
				stuspec=rsa("SPECIALITY_NAME")
				stumail=rsa("EMAIL")
				subject=rsa("THESIS_SUBJECT")
				tutorname=rsa("TUTOR_NAME")
				tutormail=rsa("TUTOR_EMAIL")
				filename="答辩安排信息"
				fieldval=Array(stuname,stuno,stuclass,stuspec,stumail,subject,tutorname,tutormail,filename)
				bSuccess=sendAnnouncementEmail(mail_id(11),stumail,fieldval)
				logtxt=logtxt&"，发送邮件给学生["&stuname&":"&stumail&"]"
				If bSuccess Then
					logtxt=logtxt&"成功。"
				Else
					logtxt=logtxt&"失败。"
				End If
				bSuccess=sendAnnouncementEmail(mail_id(12),tutormail,fieldval)
				logtxt=logtxt&"，发送邮件给导师["&tutorname&":"&tutormail&"]"
				If bSuccess Then
					logtxt=logtxt&"成功。"
				Else
					logtxt=logtxt&"失败。"
				End If
				rsa.MoveNext()
			Loop
			If numInsert+numUpdate>0 Then
				WriteLogForReviewSystem logtxt
			End If
		End If
		CloseConn conn
		addData=Array(numInsert,numUpdate)
	End Function
	
	Dim bError,errMsg
	Dim numInsert,numUpdate
	
	filename=Request.Form("filename")
	send_email=Request.Form("sendemail")="on"
	filepath=Server.MapPath("upload/xls/"&filename)
	Set connExcel=Server.CreateObject("ADODB.Connection")
	connstring="Provider=Microsoft.ACE.OLEDB.12.0;Data Source="&filepath&";Extended Properties=""Excel 12.0;HDR=Yes;IMEX=1"""
	connExcel.Open connstring
	
	Set rs=connExcel.OpenSchema(adSchemaTables)
	Do While Not rs.EOF
		If rs("TABLE_TYPE")="TABLE" Then
			table_name=rs("TABLE_NAME")
			If InStr("Sheet1$",table_name) Then Exit Do
		End If
		rs.MoveNext()
	Loop
	sql="SELECT * FROM ["&table_name&"A2:J]"
	Set rs=connExcel.Execute(sql)
	' 添加数据
	ret=addData()
	numInsert=ret(0)
	numUpdate=ret(1)
	CloseRs rs
	CloseConn connExcel
%><script type="text/javascript"><%
	If bError Then %>
	alert("导入时出错，其余<%=numInsert%>条记录已导入成功，<%=numUpdate%>条记录已更新。出错原因为：\n<%=toJsString(errMsg)%>");
<%Else %>
	alert("操作成功，<%=numInsert%>条记录已导入，<%=numUpdate%>条记录已更新。");
<%End If
%>location.href="thesisList.asp";
</script><%
End Select
%>