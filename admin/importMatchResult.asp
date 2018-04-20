<%Response.Charset="utf-8"%>
<!--#include file="../inc/ExtendedRequest.inc"-->
<!--#include file="../inc/db.asp"-->
<!--#include file="common.asp"--><%
If IsEmpty(Session("Id")) Then Response.Redirect("../error.asp?timeout")

curStep=Request.QueryString("step")
Select Case curStep
Case vbNullstring ' 文件选择页面
	Connect conn
	Set comm=Server.CreateObject("ADODB.Command")
	comm.ActiveConnection=conn
	comm.CommandText="getSemesterList"
	comm.CommandType=adCmdStoredProc
	Set semester=comm.CreateParameter("semester",adInteger,adParamInput,5,0)
	comm.Parameters.Append semester
	Set rs=comm.Execute()
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="../css/admin.css" rel="stylesheet" type="text/css" />
<script src="../scripts/jquery-1.6.3.min.js" type="text/javascript"></script>
</head>
<body bgcolor="ghostwhite">
<center><font size=4><b>导入专家匹配结果自EXCEL文件</b><br />
<form id="fmUpload" action="?step=2" method="POST" enctype="multipart/form-data">
<p>请选择要导入的 Excel 文件：<br />文件名：<input type="file" name="excelFile" size="100" /><br />
<a href="upload/matchresult_template.xlsx" target="_blank">点击下载专家匹配结果表格模板</a><br />
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
	CloseRs rs
	CloseConn conn
Case 2	' 上传进程

	Dim Upload,file,fso
	
	Set Upload=New ExtendedRequest
	Set file=Upload.File("excelFile")
	Set fso=Server.CreateObject("Scripting.FileSystemObject")
	
	' 检查上传目录是否存在
	strUploadPath = Server.MapPath("upload\xls")
	If Not fso.FolderExists(strUploadPath) Then fso.CreateFolder(strUploadPath)
	
	fileExt=LCase(file.FileExt)
	If fileExt <> "xls" And fileExt <> "xlsx" Then	' 不被允许的文件类型
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
	Set fso=Nothing
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="theme-color" content="#2D79B2" />
<title>导入专家匹配结果自EXCEL文件</title>
<link href="../css/admin.css" rel="stylesheet" type="text/css" />
<script src="../scripts/jquery-1.6.3.min.js" type="text/javascript"></script>
</head>
<body bgcolor="ghostwhite">
<center><br /><b>导入专家匹配结果自EXCEL文件</b><br /><br /><%
	If Not bError Then %>
<form id="fmUploadFinish" action="?step=3" method="POST">
<input type="hidden" name="periodid" value="<%=period_id%>" />
<input type="hidden" name="In_TASK_PROGRESS" value="<%=task_progress%>" />
<input type="hidden" name="In_REVIEW_STATUS" value="<%=review_status%>" />
<input type="hidden" name="selectmode" value="<%=select_mode%>" />
<input type="hidden" name="filename" value="<%=strDestFile%>" />
<p><%=byteFileSize%> 字节已上传，正在导入专家匹配结果...</p></form>
<script type="text/javascript">setTimeout("$('#fmUploadFinish').submit()",500);</script><%
	Else
%>
<script type="text/javascript">alert("<%=errstring%>");history.go(-1);</script><%
	End If
%></center></body></html><%
Case 3	' 数据读取，导入到数据库
	
	Dim bError,errMsg
	Dim countInsert,countUpdate,thesisIDs
	
	filename=Request.Form("filename")
	filepath=Server.MapPath("upload/xls/"&filename)
	send_email=True
	sql="CREATE TABLE #ret(CountInsert int,CountUpdate int,CountError int,FirstMatchThesisIDs nvarchar(MAX),IsError bit,ErrMsg nvarchar(MAX));"&_
			"INSERT INTO #ret EXEC importTestThesisMatchExpertResult '"&filepath&"'; SELECT * FROM #ret"
	Connect conn
	Set rs=conn.Execute(sql).NextRecordSet
	countInsert=rs("CountInsert")
	countUpdate=rs("CountUpdate")
	thesisIDs=rs("FirstMatchThesisIDs")
	bError=rs("IsError")
	errMsg=rs("ErrMsg")
	CloseRs rs
	
	If send_email And Len(thesisIDs) Then
		Dim stuname,stuno,stuclass,stuspec,stumail,subject,tutorname,tutormail,fieldval,bSuccess
		Dim logtxt:logtxt="行政人员["&Session("name")&"]匹配专家。"
		Dim mail_id:mail_id=getThesisReviewSystemMailIdByType(Now)
		' 批量发送通知邮件
		sql="SELECT STU_NAME,STU_NO,CLASS_NAME,SPECIALITY_NAME,EMAIL,THESIS_SUBJECT,TUTOR_NAME,TUTOR_EMAIL FROM VIEW_TEST_THESIS_REVIEW_INFO WHERE ID IN ("&thesisIDs&")"
		GetRecordSetNoLock conn,rs,sql,result
		Do While Not rs.EOF
			stuname=rs("STU_NAME")
			stuno=rs("STU_NO")
			stuclass=rs("CLASS_NAME")
			stuspec=rs("SPECIALITY_NAME")
			stumail=rs("EMAIL")
			subject=rs("THESIS_SUBJECT")
			tutorname=rs("TUTOR_NAME")
			tutormail=rs("TUTOR_EMAIL")
			fieldval=Array(stuname,stuno,stuclass,stuspec,stumail,subject,tutorname,tutormail)
			bSuccess=sendAnnouncementEmail(mail_id(1),stumail,fieldval)
			logtxt=logtxt&"，发送邮件给学生["&stuname&":"&stumail&"]"
			If bSuccess Then
				logtxt=logtxt&"成功。"
			Else
				logtxt=logtxt&"失败。"
			End If
			bSuccess=sendAnnouncementEmail(mail_id(2),tutormail,fieldval)
			logtxt=logtxt&"发送邮件给导师["&tutorname&":"&tutormail&"]"
			If bSuccess Then
				logtxt=logtxt&"成功。"
			Else
				logtxt=logtxt&"失败。"
			End If
			rs.MoveNext()
		Loop
		If Len(thesisIDs) Then WriteLog logtxt
		CloseRs rs
	End If
	CloseConn conn
%><script type="text/javascript"><%
	If bError Then %>
	alert("导入时出错，其余 <%=countInsert%> 条记录已导入，<%=countUpdate%> 条记录已更新。以下是错误详情：\n<%=toJsString(errMsg)%>");
<%Else %>
	alert("操作成功，<%=countInsert%> 条记录已导入，<%=countUpdate%> 条记录已更新。");
<%End If
%>location.href="thesisList.asp";
</script><%
End Select
%>