﻿<!--#include file="../inc/ExtendedRequest.inc"-->
<!--#include file="../inc/global.inc"-->
<!--#include file="common.asp"--><%
If IsEmpty(Session("Id")) Then Response.Redirect("../error.asp?timeout")

Dim conn,sql,ret,rs

step=Request.QueryString("step")
Select Case step
Case vbNullstring ' 文件选择页面
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="theme-color" content="#2D79B2" />
<title>导入教指委委员匹配结果</title>
<% useStylesheet "admin" %>
<% useScript "jquery", "upload" %>
</head>
<body>
<center><font size=4><b>导入教指委委员匹配结果</b><br />
<form id="fmUpload" action="?step=2" method="POST" enctype="multipart/form-data">
<p>请选择要导入的 Excel 文件：<br />文件名：<input type="file" name="excelFile" size="100" /><br />
<a href="upload/instruct_member_match_result_template.xlsx" target="_blank">点击下载教指委委员匹配结果表格模板</a><br />
<input type="submit" name="btnsubmit" value="提 交" />&nbsp;
<input type="button" name="btnret" value="返 回" onclick="history.go(-1)" /></p></form></center>
<script type="text/javascript">
	$(document).ready(function(){
		$('form').submit(function() {
			var valid=checkIfExcel(this.excelFile);
			if(valid) {
				$(':submit').val("正在提交，请稍候...").attr('disabled',true);
			}
			return valid;
		});
		$(':submit').attr('disabled',false);
	});
</script></body></html><%
Case 2	' 上传进程

	Dim Upload,file,fso

	Set Upload=New ExtendedRequest
	Set file=Upload.File("excelFile")
	Set fso=Server.CreateObject("Scripting.FileSystemObject")

	' 检查上传目录是否存在
	strUploadPath = Server.MapPath("upload\xls")
	If Not fso.FolderExists(strUploadPath) Then fso.CreateFolder(strUploadPath)

	file_ext=LCase(file.FileExt)
	If file_ext <> "xls" And file_ext <> "xlsx" Then	' 不被允许的文件类型
		bError = True
		errstring = "所选择的不是 Excel 文件！"
	Else
		' 生成日期格式文件名
		fileid = FormatDateTime(Now(),1)&Int(Timer)
		strDestFile = fileid&"."&file_ext
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
<title>导入教指委委员匹配结果</title>
<% useStylesheet "admin" %>
<% useScript "jquery" %>
</head>
<body>
<center><br /><b>导入教指委委员匹配结果</b><br /><br /><%
	If Not bError Then %>
<form id="fmUploadFinish" action="?step=3" method="POST">
<input type="hidden" name="selectmode" value="<%=select_mode%>" />
<input type="hidden" name="filename" value="<%=strDestFile%>" />
<p>文件上传成功，正在导入教指委委员匹配结果...</p></form>
<script type="text/javascript">setTimeout("$('#fmUploadFinish').submit()",500);</script><%
	Else
%>
<script type="text/javascript">alert("<%=errstring%>");history.go(-1);</script><%
	End If
%></center></body></html><%
Case 3	' 数据读取，导入到数据库

	Dim bError,errMsg
	Dim countInsert,countUpdate,thesis_ids

	filename=Request.Form("filename")
	filepath=Server.MapPath("upload/xls/"&filename)
	send_email=True
	sql="CREATE TABLE #ret(CountInsert int,CountUpdate int,CountError int,FirstMatchThesisIDs nvarchar(MAX),IsError bit,ErrMsg nvarchar(MAX));"&_
		"INSERT INTO #ret EXEC spImportMatchInstructMemberResult '"&filepath&"'; SELECT * FROM #ret"
	Connect conn
	Set rs=conn.Execute(sql).NextRecordSet()
	countInsert=rs("CountInsert")
	countUpdate=rs("CountUpdate")
	thesis_ids=rs("FirstMatchThesisIDs")
	bError=rs("IsError")
	errMsg=rs("ErrMsg")
	CloseRs rs

	If send_email And Len(thesis_ids) Then
		' 发送论文匹配教指委委员通知邮件
		Dim arrDissertations:arrDissertations=Split(thesis_ids,",")
		Dim activity_id,stu_type,is_sent
		Dim dict:Set dict=CreateDictionary()
		sql="SELECT ActivityId,TEACHTYPE_ID,STU_NAME,STU_NO,CLASS_NAME,SPECIALITY_NAME,EMAIL,THESIS_SUBJECT,TUTOR_NAME,TUTOR_EMAIL FROM ViewDissertations WHERE ID=?"
		For i=0 To UBound(arrDissertations)
			Set ret=ExecQuery(conn,sql,CmdParam("ID",adInteger,4,arrDissertations(i)))
			Set rs=ret("rs")
			If Not rs.EOF Then
				activity_id=rs("ActivityId")
				stu_type=rs("TEACHTYPE_ID")
				dict("stuname")=rs("STU_NAME")
				dict("stuno")=rs("STU_NO")
				dict("stuclass")=rs("CLASS_NAME")
				dict("stuspec")=rs("SPECIALITY_NAME")
				dict("stumail")=rs("EMAIL")
				dict("subject")=rs("THESIS_SUBJECT")
				dict("tutorname")=rs("TUTOR_NAME")
				dict("tutormail")=rs("TUTOR_EMAIL")
				CloseRs rs

				is_sent=sendNotifyMail(activity_id,stu_type,"lwppjzwwytzyj(xs)",dict("stumail"),dict)
				writeNotificationEventLog usertypeAdmin,Session("name"),"匹配教指委委员",usertypeStudent,_
					dict("stuname"),dict("stumail"),notifytypeMail,is_sent

				is_sent=sendNotifyMail(activity_id,stu_type,"lwppjzwwytzyj(ds)",dict("tutormail"),dict)
				writeNotificationEventLog usertypeAdmin,Session("name"),"匹配教指委委员",usertypeTutor,_
					dict("tutorname"),dict("tutormail"),notifytypeMail,is_sent
			End If
		Next
	End If
	CloseConn conn
%><script type="text/javascript"><%
	If bError Then %>
	alert("导入时出错，其余 <%=countInsert%> 条记录已导入，<%=countUpdate%> 条记录已更新。以下是错误详情：\n<%=toJsString(errMsg)%>");
<%Else %>
	alert("操作成功，<%=countInsert%> 条记录已导入，<%=countUpdate%> 条记录已更新。");
<%End If
%>location.href="paperList.asp";
</script><%
End Select
%>